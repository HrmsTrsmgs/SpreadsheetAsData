using System.Globalization;

namespace Marimo.SpreadsheetAsData.CodeGeneration;

/// <summary>
/// 生成する識別子の表記を判定し、名前の衝突や参照先の比較に使う値を求めます。
/// </summary>
static class CSharpIdentifierNames
{
    extension(string self)
    {
        /// <summary>
        /// 生成名の診断が受け付ける表記かを判定します。先頭の @ は予約語のエスケープとして扱います。
        /// </summary>
        internal bool IsValidGeneratedIdentifier
        {
            get
            {
                var start = self.StartsWith('@') ? 1 : 0;
                return start < self.Length
                    && !IsCSharpKeyword(self)
                    && Enumerable.Range(start, self.Length - start)
                        .All(index => IsAllowedCharacter(self[index], isFirst: index == start));
            }
        }

        /// <summary>
        /// 先頭のエスケープ表記と書式文字を除き、コンパイル後の識別子名を返します。
        /// </summary>
        internal string IdentifierValue =>
            string.Concat(
                from character in self.TrimStart('@')
                where char.GetUnicodeCategory(character) != UnicodeCategory.Format
                select character);
    }

    /// <summary>
    /// 生成名の診断では、連結用句読点と書式文字は先頭でも受理し、数字と結合記号は後続でのみ受理します。
    /// </summary>
    static bool IsAllowedCharacter(char character, bool isFirst) =>
        char.GetUnicodeCategory(character) switch
        {
            UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter
                or UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter
                or UnicodeCategory.OtherLetter or UnicodeCategory.LetterNumber
                or UnicodeCategory.ConnectorPunctuation or UnicodeCategory.Format => true,
            UnicodeCategory.DecimalDigitNumber or UnicodeCategory.NonSpacingMark
                or UnicodeCategory.SpacingCombiningMark => !isFirst,
            _ => false
        };

    static bool IsCSharpKeyword(string name) =>
        name is "abstract" or "as" or "base" or "bool" or "break" or "byte"
            or "case" or "catch" or "char" or "checked" or "class" or "const" or "continue"
            or "decimal" or "default" or "delegate" or "do" or "double" or "else" or "enum"
            or "event" or "explicit" or "extern" or "false" or "finally" or "fixed"
            or "float" or "for" or "foreach" or "goto" or "if" or "implicit" or "in"
            or "int" or "interface" or "internal" or "is" or "lock" or "long" or "namespace"
            or "new" or "null" or "object" or "operator" or "out" or "override" or "params"
            or "private" or "protected" or "public" or "readonly" or "ref" or "return"
            or "sbyte" or "sealed" or "short" or "sizeof" or "stackalloc" or "static"
            or "string" or "struct" or "switch" or "this" or "throw" or "true" or "try"
            or "typeof" or "uint" or "ulong" or "unchecked" or "unsafe" or "ushort" or "using"
            or "virtual" or "void" or "volatile" or "while";
}
