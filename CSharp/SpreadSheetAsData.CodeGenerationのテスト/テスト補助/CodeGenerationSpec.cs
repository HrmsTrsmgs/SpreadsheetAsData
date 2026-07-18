using System.Reflection;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// コード生成仕様テストで、生成結果を観測するための入口を提供します。
/// </summary>
static class CodeGenerationSpec
{
    /// <summary>
    /// テスト内で生成コードをコンパイルするときの既定の名前空間です。
    /// </summary>
    internal const string NamespaceName = "Generated";

    /// <summary>
    /// 指定したExcelファイルから生成されるコードを観測します。
    /// </summary>
    /// <param name="filePath">コード生成元のExcelファイル。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成コードを観測するためのテスト用オブジェクト。</returns>
    internal static GeneratedCode From(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        new(() => GenerateSources(filePath, configure));

    /// <summary>
    /// テスト内で直接用意した生成済みソースコードを観測します。
    /// </summary>
    /// <param name="sources">観測するC#ソースコード。</param>
    /// <returns>生成コードを観測するためのテスト用オブジェクト。</returns>
    internal static GeneratedCode FromSources(params string[] sources) =>
        new(() => sources);

    /// <summary>
    /// 製品コードのコード生成APIを呼び出します。
    /// </summary>
    /// <param name="filePath">コード生成元のExcelファイル。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    internal static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        WorkbookWrapperGenerator.GenerateSources(filePath, configure);

    /// <summary>
    /// 製品コードのコード生成診断APIを呼び出します。
    /// </summary>
    /// <param name="filePath">診断対象のExcelファイル。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>コード生成前に検出された診断情報。</returns>
    internal static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        WorkbookWrapperGenerator.GenerateDiagnostics(filePath, configure);

    /// <summary>
    /// 指定したExcelファイルから生成したコードをコンパイルします。
    /// </summary>
    /// <param name="filePath">コード生成元のExcelファイル。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成コードをコンパイルしたアセンブリ。</returns>
    internal static Assembly CompileGeneratedAssembly(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        GeneratedSourceCompiler.Compile(GenerateSources(filePath, configure));

    /// <summary>
    /// 生成コードをコンパイルしたアセンブリから、既定名前空間内の型を取得します。
    /// </summary>
    /// <param name="assembly">検索対象のアセンブリ。</param>
    /// <param name="typeName">既定名前空間を除いた型名。</param>
    /// <returns>指定した生成型。</returns>
    internal static Type GetRequiredType(this Assembly assembly, string typeName) =>
        assembly.GetType($"{NamespaceName}.{typeName}")
            ?? throw new InvalidOperationException(typeName);
}

/// <summary>
/// 生成されたC#ソースコード全体を、仕様テスト向けの語彙で観測します。
/// </summary>
sealed class GeneratedCode(Func<string[]> getSources)
{
    /// <summary>
    /// 生成ソースの取得を一度だけに固定するためのキャッシュです。
    /// </summary>
    string[]? sources;

    /// <summary>
    /// 観測対象のC#ソースコードを取得します。
    /// </summary>
    internal string[] Sources
    {
        get
        {
            sources ??= getSources();
            return sources;
        }
    }

    /// <summary>
    /// 生成ソースに含まれる型名を取得します。
    /// </summary>
    internal string[] TypeNames =>
        (
            from type in Sources.TypeDeclarations()
            select type.Identifier.ValueText
        ).ToArray();

    /// <summary>
    /// 指定した生成型を観測します。
    /// </summary>
    /// <param name="name">観測する生成型名。</param>
    /// <returns>指定した生成型を観測するためのテスト用オブジェクト。</returns>
    internal GeneratedType GeneratedType(string name) =>
        new(Sources.TypeDeclaration(name));
}

/// <summary>
/// 生成された1つの型を、仕様テスト向けの語彙で観測します。
/// </summary>
sealed class GeneratedType(TypeDeclarationSyntax declaration)
{
    /// <summary>
    /// 生成型の名前を取得します。
    /// </summary>
    internal string Name => declaration.Identifier.ValueText;

    /// <summary>
    /// 生成型の基底型名を取得します。
    /// </summary>
    internal string? BaseTypeName =>
        declaration.BaseList?.Types.SingleOrDefault()?.Type.ToString();

    /// <summary>
    /// 生成型に宣言されたプロパティ名を取得します。
    /// </summary>
    internal string[] PropertyNames =>
        (
            from property in declaration.Members.OfType<PropertyDeclarationSyntax>()
            select property.Identifier.ValueText
        ).ToArray();

    /// <summary>
    /// 指定した名前のプロパティ宣言を取得します。
    /// </summary>
    /// <param name="name">取得するプロパティ名。</param>
    /// <returns>指定したプロパティ宣言。</returns>
    internal PropertyDeclarationSyntax Property(string name) =>
        (
            from property in declaration.Members.OfType<PropertyDeclarationSyntax>()
            where property.Identifier.ValueText == name
            select property
        ).Single();
}
