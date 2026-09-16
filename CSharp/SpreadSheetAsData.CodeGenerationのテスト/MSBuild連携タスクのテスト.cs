using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class MSBuild連携タスクのテスト
{
    [Fact]
    public void Excel項目ごとの名前空間で同名の生成型を併用できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var orders = project.AddBasicStructureExcel(@"Orders\Master.xlsx");
        var archive = project.AddBasicStructureExcel(@"Archive\Master.xlsx");

        var tested = project.Generate([(orders, "OrdersModel"), (archive, "ArchiveModel")]);

        tested.Succeeded.Should().BeTrue();
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            tested.GeneratedFilePaths.Select(File.ReadAllText));
        generatedAssembly.GetType("OrdersModel.MasterBook").Should().NotBeNull();
        generatedAssembly.GetType("ArchiveModel.MasterBook").Should().NotBeNull();
    }

    [Fact]
    public void Excel項目の名前空間を変更または解除すると生成結果を更新します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel("BasicStructure.xlsx");
        project.Generate(excelFilePath).Succeeded.Should().BeTrue();

        var tested = project.Generate([(excelFilePath, "UpdatedModel")]);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource.Should().Contain("namespace UpdatedModel;");
        project.Generate(excelFilePath).SingleGeneratedSource.Should().Contain("namespace Generated;");
    }

    [Fact]
    public void パッケージ参照でブックごとに名前空間を分けてコンパイルし利用できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        project.AddBasicStructureExcel(@"Archive\BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellPackageReferenceSample(
            "Marimo.SpreadSheetAsData",
            """
            Invoke-MSBuild /t:Build
            Invoke-Dotnet run --project ./Consumer.csproj --no-build --no-restore
            """,
            workbookItems: """
            <SpreadsheetAsData Include="BasicStructure.xlsx" Namespace="ConsumerModel.Current" />
            <SpreadsheetAsData Include="Archive\BasicStructure.xlsx" Namespace="ConsumerModel.Archive" />
            """,
            programSource: """
            using var current = ConsumerModel.Current.BasicStructureBook.Open("BasicStructure.xlsx");
            using var archive = ConsumerModel.Archive.BasicStructureBook.Open("Archive/BasicStructure.xlsx");
            System.Console.WriteLine($"current:{current.SalesData.Name}");
            System.Console.WriteLine($"archive:{archive.SalesData.Name}");
            """);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        tested.Output.Should().Contain("current:SalesData").And.Contain("archive:SalesData");
    }

    [Theory]
    [InlineData("Marimo.SpreadSheetAsData")]
    [InlineData("Marimo.SpreadSheetAsData.Build")]
    public void パッケージ参照で生成した型をコンパイルしてExcelを読み書きできます(string packageId)
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellPackageReferenceSample(packageId,
            """
            Invoke-MSBuild /t:Build
            Invoke-Dotnet run --project ./Consumer.csproj --no-build --no-restore
            Invoke-MSBuild /t:InspectProject
            """);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        tested.Output.Should().Contain("SalesData").And.Contain("saved:Updated");
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "References.txt"))
            .Should().Contain("SpreadSheetAsData.dll").And.NotContain("SpreadSheetAsData.Build.dll");
    }

    [Theory]
    [InlineData("Marimo.SpreadSheetAsData")]
    [InlineData("Marimo.SpreadSheetAsData.Build")]
    public void パッケージ参照のデザイン時ビルドは再生成せずExcelと生成コードを紐づけます(string packageId)
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellPackageReferenceSample(packageId,
            """
            Invoke-MSBuild /t:Build
            # デザイン時にExcelを再度開いた場合は失敗する入力に置き換えます。
            Set-Content ./BasicStructure.xlsx 'Not an Excel workbook'
            Invoke-MSBuild /t:Compile /p:DesignTimeBuild=true /p:SkipCompilerExecution=true
            Invoke-MSBuild /t:InspectProject /p:DesignTimeBuild=true
            """);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().ContainSingle(it => it == "BasicStructure.SpreadsheetAsData.g.cs|BasicStructure.xlsx");
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Workbooks.txt"))
            .Should().Contain("BasicStructure.xlsx|BasicStructure.SpreadsheetAsData.g.cs");
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "OtherItems.txt"))
            .Should().NotContain(it => it.EndsWith(".SpreadsheetAsData.g.cs"));
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "AvailableItems.txt"))
            .Should().Contain("SpreadsheetAsData");
    }

    [Theory]
    [InlineData("Marimo.SpreadSheetAsData")]
    [InlineData("Marimo.SpreadSheetAsData.Build")]
    public void パッケージ利用プロジェクトのCleanは元Excelと手書きコードを残して生成ソースを削除します(string packageId)
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel("BasicStructure.xlsx");
        var original = File.ReadAllBytes(excelFilePath);
        var scriptFilePath = project.AddPowerShellPackageReferenceSample(packageId,
            """
            Invoke-MSBuild /t:Build
            if (-not (Test-Path ./BasicStructure.SpreadsheetAsData.g.cs)) { throw 'Generation did not run' }
            Invoke-MSBuild /t:Clean
            """);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.Exists(project.GeneratedFilePathFor("BasicStructure.xlsx")).Should().BeFalse();
        File.ReadAllBytes(excelFilePath).Should().Equal(original);
        File.Exists(Path.Combine(project.DirectoryPath, "Program.cs")).Should().BeTrue();
    }

    [Theory]
    [InlineData("Marimo.SpreadSheetAsData")]
    [InlineData("Marimo.SpreadSheetAsData.Build")]
    public void パッケージ参照で生成対象から外したExcelの生成コードはCompileに含めません(string packageId)
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellPackageReferenceSample(packageId,
            """
            Invoke-MSBuild /t:Build
            Invoke-MSBuild /t:InspectProject /p:IncludeWorkbook=false
            """);

        var tested = PowerShell実行結果.Run(scriptFilePath, project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.Exists(project.GeneratedFilePathFor("BasicStructure.xlsx")).Should().BeTrue();
        File.ReadAllLines(Path.Combine(project.DirectoryPath, "Compile.txt"))
            .Should().NotContain(it => it.StartsWith("BasicStructure.SpreadsheetAsData.g.cs|"));
    }

    [Fact]
    public void SpreadsheetAsData項目からExcelファイルの隣へ生成コードを出力します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.Warnings.Should().BeEmpty();
        tested.GeneratedFilePaths.Should().ContainSingle();
        tested.SingleGeneratedFilePath
            .Should().Be(project.GeneratedFilePathFor(@"Schemas\BasicStructure.xlsx"));
        tested.SingleGeneratedSource
            .Should().Contain("public partial class BasicStructureBook : Workbook");
    }

    [Fact]
    public void 生成コードは元Excelファイルへ紐づくメタデータを返します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedFile.GetMetadata("DependentUpon")
            .Should().Be("BasicStructure.xlsx");
        tested.SingleGeneratedFile.GetMetadata("DesignTimeSharedInput")
            .Should().Be("true");
    }

    [Fact]
    public void プロジェクト直下の辞書を使用します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "Sales"
            }
            """);

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public SalesSheet Sales => new(this);");
    }

    [Fact]
    public void Excelファイルと同じディレクトリの辞書をプロジェクト直下の辞書より優先します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "ProjectRootSales"
            }
            """);
        project.AddExcelDictionaryFor(
            @"Schemas\BasicStructure.xlsx",
            """
            {
              "SalesData": "SameDirectorySales"
            }
            """);

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public SameDirectorySalesSheet SameDirectorySales => new(this);");
    }

    [Fact]
    public void 異なるディレクトリにある同名Excelファイルの生成ファイルは衝突しません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var orderFilePath = project.AddBasicStructureExcel(@"Orders\Master.xlsx");
        var archiveFilePath = project.AddBasicStructureExcel(@"Archive\Master.xlsx");

        var tested = project.Generate(orderFilePath, archiveFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.GeneratedFilePaths
            .Should().BeEquivalentTo(
                [
                    project.GeneratedFilePathFor(@"Orders\Master.xlsx"),
                    project.GeneratedFilePathFor(@"Archive\Master.xlsx")
                ]);
    }

    [Fact]
    public void 辞書を追加変更削除すると生成結果を更新します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesDataSheet SalesData => new(this);");

        var dictionaryFilePath = project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "Sales"
            }
            """);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesSheet Sales => new(this);");

        File.WriteAllText(
            dictionaryFilePath,
            """
            {
              "SalesData": "Revenue"
            }
            """);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public RevenueSheet Revenue => new(this);");

        File.Delete(dictionaryFilePath);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesDataSheet SalesData => new(this);");
    }

    [Fact]
    public void 不正な辞書JSONでは辞書ファイルを示して失敗します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        var dictionaryFilePath = project.AddExcelDictionaryFor(
            @"Schemas\BasicStructure.xlsx",
            "{");

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeFalse();
        tested.Errors.Should().ContainSingle();
        tested.Errors.Single().File.Should().Be(dictionaryFilePath);
    }

    [Fact]
    public void PowerShellからdotnet_msbuildでコード生成できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellGenerationSample();

        var tested = PowerShell実行結果.Run(
            scriptFilePath,
            project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        project.GeneratedSourceFor("BasicStructure.xlsx")
            .Should().Contain("public partial class BasicStructureBook : Workbook");
    }

    [Fact]
    public void PowerShellからdotnet_msbuildのDesignTimeBuildで生成コードを参照できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        project.Generate(project.ExcelFilePathFor("BasicStructure.xlsx"));
        var scriptFilePath = project.AddPowerShellDesignTimeBuildSample();

        var tested = PowerShell実行結果.Run(
            scriptFilePath,
            project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
    }

    [Fact]
    public void SDK形式プロジェクトでは生成コードを既定Compileと重複させず元Excelファイルへ紐づけます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellSdkProjectNestingSample();

        var tested = PowerShell実行結果.Run(
            scriptFilePath,
            project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
    }

    [Fact]
    public void DesignTimeBuild用targetsはVisual_Studioが拒否する項目メタデータ条件を使いません()
    {
        MSBuild連携テストプロジェクト.BuildTargetsSource
            .Should().NotContain("%(_SpreadsheetAsDataDesignTimeGeneratedCompile.Identity)");
    }
}
