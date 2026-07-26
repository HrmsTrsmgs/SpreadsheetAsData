using System.Security.Cryptography;
using System.Text.Json;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Marimo.SpreadSheetAsData.Build;

/// <summary>
/// MSBuild の SpreadsheetAsData 項目から、型付き読み取り用C#ソースを生成します。
/// </summary>
public sealed class GenerateSpreadsheetAsData : Microsoft.Build.Utilities.Task
{
    /// <summary>
    /// コード生成対象のExcelファイルを取得または設定します。
    /// </summary>
    [Required]
    public ITaskItem[] ExcelFiles { get; set; } = [];

    /// <summary>
    /// 対象プロジェクトのディレクトリを取得または設定します。
    /// </summary>
    [Required]
    public string ProjectDirectory { get; set; } = "";

    /// <summary>
    /// 対象プロジェクトの既定名前空間を取得または設定します。
    /// </summary>
    public string RootNamespace { get; set; } = "";

    /// <summary>
    /// 生成ファイルの更新判定に使う中間出力ディレクトリを取得または設定します。
    /// </summary>
    [Required]
    public string IntermediateOutputPath { get; set; } = "";

    /// <summary>
    /// 生成したC#ソースファイルを取得します。
    /// </summary>
    [Output]
    public ITaskItem[] GeneratedFiles { get; set; } = [];

    /// <summary>
    /// 対象ExcelファイルからC#ソースを生成します。
    /// </summary>
    public override bool Execute()
    {
        var generatedFiles = new List<ITaskItem>();

        foreach (var excelFile in ExcelFiles)
        {
            Generate(excelFile, generatedFiles);
        }

        GeneratedFiles = [.. generatedFiles];

        return !Log.HasLoggedErrors;
    }

    void Generate(
        ITaskItem excelFile,
        ICollection<ITaskItem> generatedFiles)
    {
        var excelFilePath = FullPath(excelFile);
        var outputFilePath = OutputFilePath(excelFilePath);

        generatedFiles.Add(GeneratedFileItem(excelFilePath, outputFilePath));

        try
        {
            var dictionaryFilePath = DictionaryFilePath(excelFilePath);
            var fingerprint = Fingerprint(excelFilePath, dictionaryFilePath);

            var stampFilePath = StampFilePath(excelFilePath);

            if (IsUpToDate(outputFilePath, stampFilePath, fingerprint))
            {
                return;
            }

            var nameMappings = LoadNameMappings(dictionaryFilePath, excelFilePath);

            if (Log.HasLoggedErrors)
            {
                return;
            }

            var diagnostics = WorkbookWrapperGenerator.GenerateDiagnostics(
                excelFilePath,
                options => ConfigureOptions(options, nameMappings));

            foreach (var diagnostic in diagnostics)
            {
                LogDiagnostic(excelFilePath, diagnostic);
            }

            if (Log.HasLoggedErrors)
            {
                return;
            }

            var sources = WorkbookWrapperGenerator.GenerateSources(
                excelFilePath,
                options => ConfigureOptions(options, nameMappings));

            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath)!);
            Directory.CreateDirectory(Path.GetDirectoryName(stampFilePath)!);
            WriteIfChanged(outputFilePath, sources.Single());
            WriteIfChanged(stampFilePath, fingerprint);
        }
        catch (Exception exception)
        {
            Log.LogError(
                subcategory: null,
                errorCode: "SASD999",
                helpKeyword: null,
                file: excelFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: $"SpreadsheetAsData のコード生成に失敗しました。{exception.Message}");
        }
    }

    void ConfigureOptions(
        CodeGenerationOptions options,
        Dictionary<string, string> nameMappings)
    {
        options.Namespace = string.IsNullOrEmpty(RootNamespace)
            ? "Generated"
            : RootNamespace;
        options.NameMappings = nameMappings;
    }

    Dictionary<string, string> LoadNameMappings(
        string? dictionaryFilePath,
        string excelFilePath)
    {
        if (dictionaryFilePath is null)
        {
            return [];
        }

        try
        {
            return JsonSerializer.Deserialize<Dictionary<string, string>>(
                File.ReadAllText(dictionaryFilePath))
                ?? [];
        }
        catch (Exception exception) when (exception is JsonException or NotSupportedException)
        {
            Log.LogError(
                subcategory: null,
                errorCode: "SASDJSON",
                helpKeyword: null,
                file: dictionaryFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: $"SpreadsheetAsData の識別子名変換辞書を読み込めません。対象Excel: {excelFilePath}。{exception.Message}");

            return [];
        }
    }

    void LogDiagnostic(
        string excelFilePath,
        CodeGenerationDiagnostic diagnostic)
    {
        var message =
            $"SpreadsheetAsData のコード生成診断: 生成名 '{diagnostic.GeneratedName}'、元名 '{string.Join(", ", diagnostic.SourceNames)}'"
            + (diagnostic.InvalidSourceName is null
                ? ""
                : $"、無効な元名 '{diagnostic.InvalidSourceName}'");

        if (diagnostic.IsError)
        {
            Log.LogError(
                subcategory: null,
                errorCode: "SASD001",
                helpKeyword: null,
                file: excelFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: message);
        }
        else
        {
            Log.LogWarning(
                subcategory: null,
                warningCode: "SASD001",
                helpKeyword: null,
                file: excelFilePath,
                lineNumber: 0,
                columnNumber: 0,
                endLineNumber: 0,
                endColumnNumber: 0,
                message: message);
        }
    }

    string? DictionaryFilePath(string excelFilePath)
    {
        var dictionaryFileName =
            $"{Path.GetFileNameWithoutExtension(excelFilePath)}.spreadsheetasdata.json";

        var sameDirectoryDictionaryPath = Path.Combine(
            Path.GetDirectoryName(excelFilePath)!,
            dictionaryFileName);

        if (File.Exists(sameDirectoryDictionaryPath))
        {
            return sameDirectoryDictionaryPath;
        }

        var projectDirectoryDictionaryPath = Path.Combine(
            ProjectDirectory,
            dictionaryFileName);

        return File.Exists(projectDirectoryDictionaryPath)
            ? projectDirectoryDictionaryPath
            : null;
    }

    string Fingerprint(
        string excelFilePath,
        string? dictionaryFilePath) =>
        string.Join(
            Environment.NewLine,
            [
                typeof(WorkbookWrapperGenerator).Assembly.GetName().Version?.ToString() ?? "",
                RootNamespace,
                excelFilePath,
                FileFingerprint(excelFilePath),
                dictionaryFilePath ?? "",
                dictionaryFilePath is null ? "" : FileFingerprint(dictionaryFilePath)
            ]);

    static string FileFingerprint(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    static bool IsUpToDate(
        string outputFilePath,
        string stampFilePath,
        string fingerprint) =>
        File.Exists(outputFilePath)
            && File.Exists(stampFilePath)
            && File.ReadAllText(stampFilePath) == fingerprint;

    string FullPath(ITaskItem item)
    {
        var fullPath = item.GetMetadata("FullPath");
        return Path.GetFullPath(
            string.IsNullOrEmpty(fullPath)
                ? Path.Combine(ProjectDirectory, item.ItemSpec)
                : fullPath);
    }

    string OutputFilePath(string excelFilePath)
    {
        var outputDirectory = Path.GetDirectoryName(excelFilePath)!;

        return Path.Combine(
            outputDirectory,
            $"{Path.GetFileNameWithoutExtension(excelFilePath)}.SpreadsheetAsData.g.cs");
    }

    ITaskItem GeneratedFileItem(
        string excelFilePath,
        string outputFilePath)
    {
        var item = new TaskItem(outputFilePath);
        item.SetMetadata("DependentUpon", Path.GetFileName(excelFilePath));
        return item;
    }

    string StampFilePath(string excelFilePath)
    {
        var relativePath = Path.GetRelativePath(ProjectDirectory, excelFilePath);
        var relativeDirectory = Path.GetDirectoryName(relativePath);
        var stampDirectory = Path.Combine(
            FullIntermediateOutputPath,
            "SpreadsheetAsData",
            relativeDirectory ?? "");

        return Path.Combine(
            stampDirectory,
            $"{Path.GetFileNameWithoutExtension(excelFilePath)}.SpreadsheetAsData.g.cs.stamp");
    }

    string FullIntermediateOutputPath =>
        Path.GetFullPath(
            Path.IsPathRooted(IntermediateOutputPath)
                ? IntermediateOutputPath
                : Path.Combine(ProjectDirectory, IntermediateOutputPath));

    static void WriteIfChanged(
        string path,
        string content)
    {
        if (File.Exists(path) && File.ReadAllText(path) == content)
        {
            return;
        }

        File.WriteAllText(path, content);
    }
}
