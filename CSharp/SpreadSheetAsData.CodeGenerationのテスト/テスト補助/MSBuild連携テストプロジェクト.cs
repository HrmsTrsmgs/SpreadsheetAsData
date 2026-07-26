using System.Collections;
using System.Diagnostics;
using FluentAssertions;
using Marimo.SpreadSheetAsData.Build;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// MSBuild連携のテストで、利用者プロジェクトに近い一時ディレクトリを扱います。
/// </summary>
sealed class MSBuild連携テストプロジェクト : IDisposable
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string Configuration = "Debug";
    const string TargetFramework = "net10.0";

    MSBuild連携テストプロジェクト()
    {
        DirectoryPath = Path.Combine(
            Path.GetTempPath(),
            "SpreadsheetAsData.Tests",
            Guid.NewGuid().ToString("N"));
    }

    /// <summary>
    /// テスト対象プロジェクトのルートディレクトリです。
    /// </summary>
    internal string DirectoryPath { get; }

    /// <summary>
    /// 新しい一時プロジェクトを作成します。
    /// </summary>
    internal static MSBuild連携テストプロジェクト Create() =>
        new();

    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
        {
            Directory.Delete(DirectoryPath, true);
        }
    }

    /// <summary>
    /// 既存の基本構造テストブックを、指定した相対パスへ追加します。
    /// </summary>
    /// <param name="relativePath">一時プロジェクト内のExcelファイル相対パス。</param>
    /// <returns>追加したExcelファイルの絶対パス。</returns>
    internal string AddBasicStructureExcel(string relativePath)
    {
        var destinationPath = Path.Combine(DirectoryPath, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        File.Copy(BasicStructureExcelFilePath, destinationPath);
        return destinationPath;
    }

    /// <summary>
    /// プロジェクト直下へ、指定したExcelファイル用の識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="excelFileName">対応するExcelファイル名。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddProjectDictionaryFor(
        string excelFileName,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                DictionaryFileName(excelFileName)),
            json);

    /// <summary>
    /// Excelファイルと同じディレクトリへ、識別子名変換辞書を置きます。
    /// </summary>
    /// <param name="excelRelativePath">対応するExcelファイルの一時プロジェクト内相対パス。</param>
    /// <param name="json">辞書JSON。</param>
    /// <returns>作成した辞書ファイルの絶対パス。</returns>
    internal string AddExcelDictionaryFor(
        string excelRelativePath,
        string json) =>
        WriteDictionary(
            Path.Combine(
                DirectoryPath,
                Path.GetDirectoryName(excelRelativePath) ?? "",
                DictionaryFileName(excelRelativePath)),
            json);

    /// <summary>
    /// MSBuildタスクを直接実行します。
    /// </summary>
    /// <param name="excelFilePaths">生成対象Excelファイルの絶対パス。</param>
    /// <returns>タスクの実行結果。</returns>
    internal MSBuild連携タスク実行結果 Generate(params string[] excelFilePaths)
    {
        var buildEngine = new RecordingBuildEngine();
        var task = new GenerateSpreadsheetAsData
        {
            BuildEngine = buildEngine,
            ExcelFiles = [.. excelFilePaths.Select(it => new TaskItem(it))],
            ProjectDirectory = DirectoryPath,
            RootNamespace = "Generated",
            IntermediateOutputPath = Path.Combine("obj", Configuration, TargetFramework)
        };

        return new(
            task.Execute(),
            task.GeneratedFiles,
            buildEngine.Errors,
            buildEngine.Warnings);
    }

    /// <summary>
    /// 指定したExcelファイルに対応する生成ファイルの絶対パスを取得します。
    /// </summary>
    /// <param name="excelRelativePath">一時プロジェクト内のExcelファイル相対パス。</param>
    /// <returns>生成ファイルの絶対パス。</returns>
    internal string GeneratedFilePathFor(string excelRelativePath) =>
        Path.Combine(
            DirectoryPath,
            Path.GetDirectoryName(excelRelativePath) ?? "",
            $"{Path.GetFileNameWithoutExtension(excelRelativePath)}.SpreadsheetAsData.g.cs");

    /// <summary>
    /// 指定したExcelファイルに対応する生成ソースを読み込みます。
    /// </summary>
    /// <param name="excelRelativePath">一時プロジェクト内のExcelファイル相対パス。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    internal string GeneratedSourceFor(string excelRelativePath) =>
        File.ReadAllText(GeneratedFilePathFor(excelRelativePath));

    /// <summary>
    /// PowerShellからdotnet msbuildを起動するサンプルを作成します。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトの絶対パス。</returns>
    internal string AddPowerShellGenerationSample()
    {
        AddPackageLayout();
        File.WriteAllText(
            Path.Combine(DirectoryPath, "SpreadsheetAsData.Generate.proj"),
            """
            <Project>
              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.props" />

              <PropertyGroup>
                <RootNamespace>Generated</RootNamespace>
                <IntermediateOutputPath>obj\Debug\net10.0\</IntermediateOutputPath>
                <DesignTimeBuild>false</DesignTimeBuild>
              </PropertyGroup>

              <ItemGroup>
                <SpreadsheetAsData Include="BasicStructure.xlsx" />
              </ItemGroup>

              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.targets" />

              <Target Name="Build" DependsOnTargets="GenerateSpreadsheetAsDataSources" />
            </Project>
            """);

        var scriptFilePath = Path.Combine(DirectoryPath, "Generate.ps1");
        File.WriteAllText(
            scriptFilePath,
            """
            $ErrorActionPreference = 'Stop'

            dotnet msbuild .\SpreadsheetAsData.Generate.proj /t:Build /nologo /v:minimal

            $generatedFile = '.\BasicStructure.SpreadsheetAsData.g.cs'

            if (-not (Test-Path $generatedFile)) {
                throw "Generated file was not created: $generatedFile"
            }

            if (-not (Select-String -Path $generatedFile -SimpleMatch 'public partial class BasicStructureBook : Workbook')) {
                throw "Generated source does not contain BasicStructureBook."
            }
            """);

        return scriptFilePath;
    }

    void AddPackageLayout()
    {
        var buildTransitiveDirectory = Path.Combine(DirectoryPath, "Package", "buildTransitive");
        var toolsDirectory = Path.Combine(DirectoryPath, "Package", "tools", TargetFramework);
        Directory.CreateDirectory(buildTransitiveDirectory);
        Directory.CreateDirectory(toolsDirectory);

        CopyRepositoryFile(
            @"SpreadSheetAsData.Build\buildTransitive\Marimo.SpreadSheetAsData.Build.props",
            Path.Combine(buildTransitiveDirectory, "Marimo.SpreadSheetAsData.Build.props"));
        CopyRepositoryFile(
            @"SpreadSheetAsData.Build\buildTransitive\Marimo.SpreadSheetAsData.Build.targets",
            Path.Combine(buildTransitiveDirectory, "Marimo.SpreadSheetAsData.Build.targets"));

        foreach (var filePath in Directory.GetFiles(AppContext.BaseDirectory))
        {
            if (Path.GetFileName(filePath) is
                "SpreadSheetAsData.Build.dll"
                or "SpreadSheetAsData.Build.deps.json"
                or "SpreadSheetAsData.CodeGeneration.dll"
                or "SpreadSheetAsData.dll"
                or "DocumentFormat.OpenXml.dll"
                or "System.Interactive.dll"
                or "System.IO.Packaging.dll")
            {
                File.Copy(
                    filePath,
                    Path.Combine(toolsDirectory, Path.GetFileName(filePath)),
                    true);
            }
        }
    }

    static void CopyRepositoryFile(
        string relativePathFromCSharpDirectory,
        string destinationPath)
    {
        File.Copy(
            Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    relativePathFromCSharpDirectory)),
            destinationPath,
            true);
    }

    static string WriteDictionary(
        string filePath,
        string json)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
        File.WriteAllText(filePath, json);
        return filePath;
    }

    static string DictionaryFileName(string excelFileName) =>
        $"{Path.GetFileNameWithoutExtension(excelFileName)}.spreadsheetasdata.json";
}

/// <summary>
/// MSBuild連携タスクの実行結果を、テスト内で観測しやすい形にまとめます。
/// </summary>
sealed record MSBuild連携タスク実行結果(
    bool Succeeded,
    ITaskItem[] GeneratedFiles,
    IReadOnlyList<BuildErrorEventArgs> Errors,
    IReadOnlyList<BuildWarningEventArgs> Warnings)
{
    /// <summary>
    /// 生成ファイルの絶対パス一覧です。
    /// </summary>
    internal string[] GeneratedFilePaths =>
        [.. GeneratedFiles.Select(it => it.ItemSpec)];

    /// <summary>
    /// 生成ファイルが1つであるテストで、そのタスク項目を取得します。
    /// </summary>
    internal ITaskItem SingleGeneratedFile =>
        GeneratedFiles.Single();

    /// <summary>
    /// 生成ファイルが1つであるテストで、そのファイルパスを取得します。
    /// </summary>
    internal string SingleGeneratedFilePath =>
        GeneratedFilePaths.Single();

    /// <summary>
    /// 生成ファイルが1つであるテストで、そのC#ソースを取得します。
    /// </summary>
    internal string SingleGeneratedSource =>
        File.ReadAllText(SingleGeneratedFilePath);
}

/// <summary>
/// PowerShellからの起動結果を表します。
/// </summary>
sealed record PowerShell実行結果(
    int ExitCode,
    string Output)
{
    /// <summary>
    /// PowerShellスクリプトを実行し、標準出力と標準エラーをまとめて取得します。
    /// </summary>
    /// <param name="scriptFilePath">実行するPowerShellスクリプト。</param>
    /// <param name="workingDirectory">実行時の作業ディレクトリ。</param>
    /// <returns>PowerShellの終了コードと出力。</returns>
    internal static PowerShell実行結果 Run(
        string scriptFilePath,
        string workingDirectory)
    {
        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptFilePath}\"",
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        }) ?? throw new InvalidOperationException("PowerShellを起動できませんでした。");

        var output = process.StandardOutput.ReadToEnd()
            + process.StandardError.ReadToEnd();
        process.WaitForExit();

        return new(process.ExitCode, output);
    }
}

/// <summary>
/// MSBuildタスクが出力したエラーと警告を記録します。
/// </summary>
sealed class RecordingBuildEngine : IBuildEngine
{
    internal List<BuildErrorEventArgs> Errors { get; } = [];

    internal List<BuildWarningEventArgs> Warnings { get; } = [];

    public bool ContinueOnError => false;

    public int LineNumberOfTaskNode => 0;

    public int ColumnNumberOfTaskNode => 0;

    public string ProjectFileOfTaskNode => "";

    public bool BuildProjectFile(
        string projectFileName,
        string[] targetNames,
        IDictionary globalProperties,
        IDictionary targetOutputs) =>
        throw new NotImplementedException();

    public void LogCustomEvent(CustomBuildEventArgs e)
    {
    }

    public void LogErrorEvent(BuildErrorEventArgs e)
    {
        Errors.Add(e);
    }

    public void LogMessageEvent(BuildMessageEventArgs e)
    {
    }

    public void LogWarningEvent(BuildWarningEventArgs e)
    {
        Warnings.Add(e);
    }
}
