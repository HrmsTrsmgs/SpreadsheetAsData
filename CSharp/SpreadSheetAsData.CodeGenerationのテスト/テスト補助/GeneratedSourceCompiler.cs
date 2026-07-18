using System.Reflection;
using System.Runtime.Loader;
using Marimo.SpreadSheetAsData;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

static class GeneratedSourceCompiler
{
    internal static Assembly Compile(IEnumerable<string> sources)
    {
        var syntaxTrees =
            from source in sources
            select CSharpSyntaxTree.ParseText(source);

        var compilation = CSharpCompilation.Create(
            $"SpreadsheetAsData.Generated.{Guid.NewGuid():N}",
            syntaxTrees,
            References,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

        using var stream = new MemoryStream();
        var result = compilation.Emit(stream);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                string.Join(
                    Environment.NewLine,
                    from diagnostic in result.Diagnostics
                    where diagnostic.Severity == DiagnosticSeverity.Error
                    select diagnostic.ToString()));
        }

        stream.Position = 0;
        return AssemblyLoadContext.Default.LoadFromStream(stream);
    }

    static IEnumerable<MetadataReference> References =>
        (
            from path in TrustedPlatformAssemblyPaths
                .Append(typeof(Workbook).Assembly.Location)
            where !string.IsNullOrEmpty(path)
            group path by path into paths
            select MetadataReference.CreateFromFile(paths.Key)
        );

    static IEnumerable<string> TrustedPlatformAssemblyPaths =>
        ((string?)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES"))
            ?.Split(Path.PathSeparator)
            ?? [];
}


