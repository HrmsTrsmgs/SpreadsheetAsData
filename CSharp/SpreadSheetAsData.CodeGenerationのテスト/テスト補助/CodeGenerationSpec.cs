using System.Reflection;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

static class CodeGenerationSpec
{
    internal const string NamespaceName = "Generated";

    internal static GeneratedCode From(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        new(() => GenerateSources(filePath, configure));

    internal static GeneratedCode FromSources(params string[] sources) =>
        new(() => sources);

    internal static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        WorkbookWrapperGenerator.GenerateSources(filePath, configure);

    internal static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        WorkbookWrapperGenerator.GenerateDiagnostics(filePath, configure);

    internal static Assembly CompileGeneratedAssembly(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        GeneratedSourceCompiler.Compile(GenerateSources(filePath, configure));

    internal static Type GetRequiredType(this Assembly assembly, string typeName) =>
        assembly.GetType($"{NamespaceName}.{typeName}")
            ?? throw new InvalidOperationException(typeName);
}

sealed class GeneratedCode(Func<string[]> getSources)
{
    string[]? sources;

    internal string[] Sources
    {
        get
        {
            sources ??= getSources();
            return sources;
        }
    }

    internal string[] TypeNames =>
        (
            from type in Sources.TypeDeclarations()
            select type.Identifier.ValueText
        ).ToArray();

    internal GeneratedType GeneratedType(string name) =>
        new(Sources.TypeDeclaration(name));
}

sealed class GeneratedType(TypeDeclarationSyntax declaration)
{
    internal string Name => declaration.Identifier.ValueText;

    internal string? BaseTypeName =>
        declaration.BaseList?.Types.SingleOrDefault()?.Type.ToString();

    internal string[] PropertyNames =>
        (
            from property in declaration.Members.OfType<PropertyDeclarationSyntax>()
            select property.Identifier.ValueText
        ).ToArray();

    internal PropertyDeclarationSyntax Property(string name) =>
        (
            from property in declaration.Members.OfType<PropertyDeclarationSyntax>()
            where property.Identifier.ValueText == name
            select property
        ).Single();
}
