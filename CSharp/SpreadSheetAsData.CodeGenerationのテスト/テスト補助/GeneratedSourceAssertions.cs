using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// 生成ソースの構文木から宣言やコメントを取り出す低レベルなテスト補助です。
/// </summary>
static class GeneratedSourceAssertions
{
    extension(IEnumerable<string> self)
    {
        /// <summary>
        /// 生成ソースに含まれる型宣言を取得します。
        /// </summary>
        /// <returns>ソース内の型宣言。</returns>
        internal IEnumerable<TypeDeclarationSyntax> TypeDeclarations() =>
            from source in self
            from type in
                CSharpSyntaxTree
                    .ParseText(source)
                    .GetCompilationUnitRoot()
                    .DescendantNodes()
                    .OfType<TypeDeclarationSyntax>()
            select type;

        /// <summary>
        /// 生成ソースから指定した名前の型宣言を取得します。
        /// </summary>
        /// <param name="name">取得する型名。</param>
        /// <returns>指定した型宣言。</returns>
        internal TypeDeclarationSyntax TypeDeclaration(string name) =>
            (
                from type in self.TypeDeclarations()
                where type.Identifier.ValueText == name
                select type
            ).Single();
    }

    extension(TypeDeclarationSyntax self)
    {
        /// <summary>
        /// 指定した型宣言から指定した名前のプロパティ宣言を取得します。
        /// </summary>
        /// <param name="propertyName">取得するプロパティ名。</param>
        /// <returns>指定したプロパティ宣言。</returns>
        internal PropertyDeclarationSyntax PropertyDeclaration(string propertyName) =>
            (
                from property in
                    self.Members
                        .OfType<PropertyDeclarationSyntax>()
                where property.Identifier.ValueText == propertyName
                select property
            ).Single();
    }

    extension(MemberDeclarationSyntax self)
    {
        /// <summary>
        /// 宣言に付けられたXMLコメントの summary 本文を取得します。
        /// </summary>
        /// <returns>summary 本文。summary がない場合は null。</returns>
        internal string? SummaryText()
        {
            var trivia = self.GetLeadingTrivia()
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
                    from token in
                        summary.Content
                            .OfType<XmlTextSyntax>()
                            .SelectMany(it => it.TextTokens)
                    let text = token.ValueText.Trim()
                    where text != ""
                    select text);
        }
    }

    extension(PropertyDeclarationSyntax self)
    {
        /// <summary>
        /// 指定した属性に渡された引数を文字列として取得します。
        /// </summary>
        /// <param name="attributeName">取得する属性名。</param>
        /// <returns>属性に指定された引数。</returns>
        internal IEnumerable<string> AttributeArguments(string attributeName) =>
            from attribute in self.AttributeLists.SelectMany(it => it.Attributes)
            where attribute.Name.ToString() == attributeName
                || attribute.Name.ToString() == attributeName.Replace("Attribute", "")
            from argument in
                attribute.ArgumentList?.Arguments
                    ?? []
            select argument.ToString().Trim('"');
    }
}
