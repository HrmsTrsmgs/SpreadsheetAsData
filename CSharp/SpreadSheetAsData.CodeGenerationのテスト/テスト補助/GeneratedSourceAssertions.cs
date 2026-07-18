using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

static class GeneratedSourceAssertions
{
    internal static IEnumerable<TypeDeclarationSyntax> TypeDeclarations(
        this IEnumerable<string> sources) =>
        from source in sources
        from type in CSharpSyntaxTree
            .ParseText(source)
            .GetCompilationUnitRoot()
            .DescendantNodes()
            .OfType<TypeDeclarationSyntax>()
        select type;

    internal static TypeDeclarationSyntax TypeDeclaration(
        this IEnumerable<string> sources,
        string name) =>
        (
            from type in sources.TypeDeclarations()
            where type.Identifier.ValueText == name
            select type
        ).Single();

    internal static PropertyDeclarationSyntax PropertyDeclaration(
        this IEnumerable<string> sources,
        string typeName,
        string propertyName) =>
        (
            from property in sources
                .TypeDeclaration(typeName)
                .Members
                .OfType<PropertyDeclarationSyntax>()
            where property.Identifier.ValueText == propertyName
            select property
        ).Single();

    internal static string? SummaryText(this MemberDeclarationSyntax declaration)
    {
        var trivia = declaration.GetLeadingTrivia()
            .Select(it => it.GetStructure())
            .OfType<DocumentationCommentTriviaSyntax>()
            .SingleOrDefault();

        var summary = trivia
            ?.Content
            .OfType<XmlElementSyntax>()
            .SingleOrDefault(it => it.StartTag.Name.LocalName.ValueText == "summary");

        return summary == null
            ? null
            : string.Join(
                " ",
                from token in summary.Content
                    .OfType<XmlTextSyntax>()
                    .SelectMany(it => it.TextTokens)
                let text = token.ValueText.Trim()
                where text != ""
                select text);
    }

    internal static IEnumerable<string> AttributeArguments(
        this PropertyDeclarationSyntax property,
        string attributeName) =>
        from attribute in property.AttributeLists.SelectMany(it => it.Attributes)
        where attribute.Name.ToString() == attributeName
            || attribute.Name.ToString() == attributeName.Replace("Attribute", "")
        from argument in attribute.ArgumentList?.Arguments
            ?? []
        select argument.ToString().Trim('"');
}


