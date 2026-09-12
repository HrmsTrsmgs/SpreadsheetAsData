using System.Collections;
using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using System.Xml.Linq;
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

    /// <summary>
    /// パッケージへ含めるMSBuild targetsの内容です。
    /// </summary>
    internal static string BuildTargetsSource =>
        File.ReadAllText(RepositoryFilePath(
            @"SpreadSheetAsData.Build\buildTransitive\Marimo.SpreadSheetAsData.Build.targets"));

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
    /// 指定したExcelファイルの絶対パスを取得します。
    /// </summary>
    /// <param name="excelRelativePath">一時プロジェクト内のExcelファイル相対パス。</param>
    /// <returns>Excelファイルの絶対パス。</returns>
    internal string ExcelFilePathFor(string excelRelativePath) =>
        Path.Combine(DirectoryPath, excelRelativePath);

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

    /// <summary>
    /// PowerShellからdotnet msbuildのDesignTimeBuildを起動するサンプルを作成します。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトの絶対パス。</returns>
    internal string AddPowerShellDesignTimeBuildSample()
    {
        AddPackageLayout();
        File.WriteAllText(
            Path.Combine(DirectoryPath, "SpreadsheetAsData.DesignTimeBuild.proj"),
            """
            <Project>
              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.props" />

              <PropertyGroup>
                <RootNamespace>Generated</RootNamespace>
                <IntermediateOutputPath>obj\Debug\net10.0\</IntermediateOutputPath>
                <DesignTimeBuild>true</DesignTimeBuild>
              </PropertyGroup>

              <ItemGroup>
                <SpreadsheetAsData Include="BasicStructure.xlsx" />
              </ItemGroup>

              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.targets" />

              <Target Name="Build">
                <Error
                  Condition="'@(Compile)' == ''"
                  Text="Generated source was not added to Compile." />
              </Target>
            </Project>
            """);

        var scriptFilePath = Path.Combine(DirectoryPath, "DesignTimeBuild.ps1");
        File.WriteAllText(
            scriptFilePath,
            """
            $ErrorActionPreference = 'Stop'

            dotnet msbuild .\SpreadsheetAsData.DesignTimeBuild.proj /t:Build /nologo /v:minimal
            """);

        return scriptFilePath;
    }

    /// <summary>
    /// SDK形式プロジェクトで、生成コードの重複登録とVisual Studio向けメタデータを確認するサンプルを作成します。
    /// </summary>
    /// <returns>実行するPowerShellスクリプトの絶対パス。</returns>
    internal string AddPowerShellSdkProjectNestingSample()
    {
        AddPackageLayout();
        File.WriteAllText(
            Path.Combine(DirectoryPath, "SpreadsheetAsData.SdkProject.csproj"),
            """
            <Project Sdk="Microsoft.NET.Sdk">
              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.props" />

              <PropertyGroup>
                <OutputType>Library</OutputType>
                <TargetFramework>net10.0</TargetFramework>
                <ImplicitUsings>enable</ImplicitUsings>
                <Nullable>enable</Nullable>
                <RootNamespace>Generated</RootNamespace>
                <RestoreProjectStyle>None</RestoreProjectStyle>
              </PropertyGroup>

              <ItemGroup>
                <Reference Include="SpreadSheetAsData" HintPath="Package\tools\net10.0\SpreadSheetAsData.dll" />
                <SpreadsheetAsData Include="BasicStructure.xlsx" />
              </ItemGroup>

              <Import Project="Package\buildTransitive\Marimo.SpreadSheetAsData.Build.targets" />

              <Target Name="CheckGeneratedCompileMetadata">
                <ItemGroup>
                  <_SpreadsheetAsDataExpectedGeneratedCompile
                    Include="@(Compile)"
                    Condition="'%(Compile.Filename)%(Compile.Extension)' == 'BasicStructure.SpreadsheetAsData.g.cs' and '%(Compile.DependentUpon)' == 'BasicStructure.xlsx'" />
                  <_SpreadsheetAsDataVisibleGeneratedNonCompile
                    Include="@(None);@(Content)"
                    Condition="'%(Filename)%(Extension)' == 'BasicStructure.SpreadsheetAsData.g.cs'" />
                  <_SpreadsheetAsDataParentWithLastGeneratedOutput
                    Include="@(SpreadsheetAsData)"
                    Condition="'%(SpreadsheetAsData.Filename)%(SpreadsheetAsData.Extension)' == 'BasicStructure.xlsx' and '%(SpreadsheetAsData.LastGenOutput)' == 'BasicStructure.SpreadsheetAsData.g.cs'" />
                </ItemGroup>

                <Error
                  Condition="'%(Compile.Filename)%(Compile.Extension)' == 'BasicStructure.SpreadsheetAsData.g.cs' and '%(Compile.DependentUpon)' != 'BasicStructure.xlsx'"
                  Text="Generated source was not nested under the Excel file." />
                <Error
                  Condition="'@(_SpreadsheetAsDataExpectedGeneratedCompile)' == ''"
                  Text="Generated source was not declared as Compile with Excel nesting metadata." />
                <Error
                  Condition="'@(_SpreadsheetAsDataVisibleGeneratedNonCompile)' != ''"
                  Text="Generated source was also visible as non-Compile item." />
                <Error
                  Condition="'@(_SpreadsheetAsDataParentWithLastGeneratedOutput)' == ''"
                  Text="Excel file did not declare the generated source as LastGenOutput." />
              </Target>
            </Project>
            """);

        var scriptFilePath = Path.Combine(DirectoryPath, "SdkProject.ps1");
        File.WriteAllText(
            scriptFilePath,
            """
            $ErrorActionPreference = 'Stop'

            function Invoke-DotnetMSBuild {
                dotnet msbuild @args
                if ($LASTEXITCODE -ne 0) {
                    throw "dotnet msbuild $($args -join ' ') failed with exit code $LASTEXITCODE."
                }
            }

            Invoke-DotnetMSBuild .\SpreadsheetAsData.SdkProject.csproj /t:CheckGeneratedCompileMetadata /nologo /v:minimal
            Invoke-DotnetMSBuild .\SpreadsheetAsData.SdkProject.csproj /t:GenerateSpreadsheetAsDataSources /nologo /v:minimal
            Invoke-DotnetMSBuild .\SpreadsheetAsData.SdkProject.csproj /t:CheckGeneratedCompileMetadata /nologo /v:minimal /p:DesignTimeBuild=true
            """);

        return scriptFilePath;
    }

    /// <summary>
    /// 実パッケージから復元する利用者プロジェクトと、指定された検証コマンドを実行するスクリプトを作ります。
    /// 外部依存は復元済みキャッシュから取得し、今回のパッケージと復元先をテストごとに分離します。
    /// </summary>
    /// <param name="packageId">全部入りまたはBuild単体のパッケージID。</param>
    /// <param name="commands">復元後に実行するPowerShellコマンド。</param>
    /// <returns>pack、restore、検証の順で実行するスクリプトのパス。</returns>
    internal string AddPowerShellPackageReferenceSample(string packageId, string commands)
    {
        Directory.CreateDirectory(DirectoryPath);
        using var assets = JsonDocument.Parse(File.ReadAllText(
            RepositoryFilePath(@"SpreadSheetAsData.Build\obj\project.assets.json")));
        File.WriteAllText(
            Path.Combine(DirectoryPath, "NuGet.Config"),
            new XElement("configuration",
                new XElement("packageSources",
                    new XElement("clear"),
                    new XElement("add", new XAttribute("key", "local"), new XAttribute("value", "packages")),
                    assets.RootElement.GetProperty("packageFolders").EnumerateObject().Select((it, index) =>
                        new XElement("add", new XAttribute("key", $"cache{index}"), new XAttribute("value", it.Name)))),
                new XElement("packageSourceMapping",
                    new XElement("packageSource", new XAttribute("key", "local"),
                        new XElement("package", new XAttribute("pattern", "Marimo.SpreadSheetAsData*"))),
                    assets.RootElement.GetProperty("packageFolders").EnumerateObject().Select((it, index) =>
                        new XElement("packageSource", new XAttribute("key", $"cache{index}"),
                            new XElement("package", new XAttribute("pattern", "*"))))))
                .ToString());
        var version = XDocument.Load(RepositoryFilePath(@"SpreadSheetAsData.Package\SpreadSheetAsData.Package.csproj"))
            .Descendants("Version").Single().Value;
        File.WriteAllText(
            Path.Combine(DirectoryPath, "Consumer.csproj"),
            $$"""
            <Project Sdk="Microsoft.NET.Sdk">
              <PropertyGroup>
                <TargetFramework>net10.0</TargetFramework>
                <OutputType>Exe</OutputType>
                <RootNamespace>ConsumerModel</RootNamespace>
                <UseSharedCompilation>false</UseSharedCompilation>
                <RestorePackagesPath>$(MSBuildProjectDirectory)/restored</RestorePackagesPath>
                <!-- テスト専用のオフライン復元。通常の利用者の監査設定は変更しません。 -->
                <NuGetAudit>false</NuGetAudit>
                <IncludeWorkbook Condition="'$(IncludeWorkbook)' == ''">true</IncludeWorkbook>
              </PropertyGroup>
              <ItemGroup>
                <PackageReference Include="{{packageId}}" Version="{{version}}" />
                <SpreadsheetAsData Include="BasicStructure.xlsx" Condition="'$(IncludeWorkbook)' == 'true'" />
              </ItemGroup>
              <Target Name="InspectProject" DependsOnTargets="ResolveReferences">
                <WriteLinesToFile File="Compile.txt" Lines="@(Compile->'%(Filename)%(Extension)|%(DependentUpon)')" Overwrite="true" />
                <WriteLinesToFile File="OtherItems.txt" Lines="@(None);@(Content)" Overwrite="true" />
                <WriteLinesToFile File="Workbooks.txt" Lines="@(SpreadsheetAsData->'%(Filename)%(Extension)|%(LastGenOutput)')" Overwrite="true" />
                <WriteLinesToFile File="AvailableItems.txt" Lines="@(AvailableItemName)" Overwrite="true" />
                <WriteLinesToFile File="References.txt" Lines="@(ReferencePath->'%(Filename)%(Extension)')" Overwrite="true" />
              </Target>
            </Project>
            """);
        File.WriteAllText(
            Path.Combine(DirectoryPath, "Program.cs"),
            """
            using ConsumerModel;

            using (var book = BasicStructureBook.Open("BasicStructure.xlsx"))
            {
                System.Console.WriteLine(book.SalesData.Name);
                book.SalesData.Cells["A1"].Value = "Updated";
                book.SaveAs("Updated.xlsx");
            }
            using var saved = BasicStructureBook.Open("Updated.xlsx");
            string value = saved.SalesData.Cells["A1"].Value;
            System.Console.WriteLine($"saved:{value}");
            """);
        var configuration = typeof(GenerateSpreadsheetAsData).Assembly
            .GetCustomAttributes<AssemblyConfigurationAttribute>().Single().Configuration;
        var scriptFilePath = Path.Combine(DirectoryPath, "PackageReference.ps1");
        File.WriteAllText(
            scriptFilePath,
            $$"""
            $ErrorActionPreference = 'Stop'
            [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            function Invoke-Dotnet {
                dotnet @args
                if ($LASTEXITCODE -ne 0) { throw "dotnet failed: $LASTEXITCODE" }
            }
            # 任意指定のVisual Studio版MSBuildでも、同じ利用プロジェクトを検証できます。
            function Invoke-MSBuild {
                if ($env:SPREADSHEETASDATA_TEST_MSBUILD) {
                    & $env:SPREADSHEETASDATA_TEST_MSBUILD ./Consumer.csproj /nr:false /nologo /v:minimal @args
                } else {
                    dotnet msbuild ./Consumer.csproj /nr:false /nologo /v:minimal @args
                }
                if ($LASTEXITCODE -ne 0) { throw "MSBuild failed: $LASTEXITCODE" }
            }
            foreach ($projectName in @('SpreadSheetAsData', 'SpreadSheetAsData.CodeGeneration', 'SpreadSheetAsData.Build', 'SpreadSheetAsData.Package')) {
                $projectPath = Join-Path '{{RepositoryFilePath("").Replace("'", "''")}}' "$projectName/$projectName.csproj"
                Invoke-Dotnet pack $projectPath --no-build --no-restore --disable-build-servers --configuration '{{configuration}}' --output ./packages --verbosity quiet '-p:TreatWarningsAsErrors=true'
            }
            Invoke-Dotnet restore ./Consumer.csproj --configfile ./NuGet.Config --verbosity quiet
            {{commands}}
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
            RepositoryFilePath(relativePathFromCSharpDirectory),
            destinationPath,
            true);
    }

    static string RepositoryFilePath(string relativePathFromCSharpDirectory) =>
        Path.GetFullPath(
            Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                relativePathFromCSharpDirectory));

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
