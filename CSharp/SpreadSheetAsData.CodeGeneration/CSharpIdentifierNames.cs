using System.Globalization;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// 生成された識別子の表記を、名前の比較に使用する形へ揃えます。
/// </summary>
static class CSharpIdentifierNames
{
    extension(string self)
    {
        /// <summary>
        /// 先頭のエスケープ表記と書式文字を除き、コンパイル後の識別子名を返します。
        /// </summary>
        internal string IdentifierValue =>
            string.Concat(
                from character in self.TrimStart('@')
                where char.GetUnicodeCategory(character) != UnicodeCategory.Format
                select character);
    }
}
