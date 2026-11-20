using System.Globalization;

using Marimo.SpreadsheetAsData;

namespace Marimo.SpreadsheetAsData.CodeGeneration;

/// <summary>
/// 生成する型とメンバーの名前について、無効名と名前衝突を診断します。
/// </summary>
static class GeneratedNameDiagnostics
{
    /// <summary>
    /// 開かれたブックについて、生成先ごとの診断を既定の順序で返します。
    /// </summary>
    internal static CodeGenerationDiagnostic[] Diagnose(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        [
            .. InvalidPropertyNameDiagnostics(
                from sheet in book.Sheets.Values
                select sheet.Name,
                options),
            .. InvalidPropertyNameDiagnostics(
                from table in book.Tables
                select table.Name,
                options),
            .. from member in BookDefinedNames(book, options)
               where IsInvalidPropertyName(member.Name)
               select new CodeGenerationDiagnostic(true, member.Name, [member.SourceName], member.SourceName),
            .. TypeNameDiagnostics(filePath, book, options),
            .. BookPropertyNameDiagnostics(filePath, book, options),
            .. SheetDefinedNameDiagnostics(book, options),
            .. SheetDefinedNameTableDiagnostics(book, options),
            .. SheetTableNameDiagnostics(book, options),
            .. DataPropertyNameDiagnostics(filePath, book, options),
            .. from table in book.Tables
               from diagnostic in ColumnPropertyNameDiagnostics(table, options)
               select diagnostic
        ];

    /// <summary>
    /// ブック・ワークシート・テーブルから生成する型名同士の衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> TypeNameDiagnostics(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        NameCollisionDiagnostics(
            GroupMemberNames(
                [
                    .. from suffix in new[] { "Book", "Data" }
                       select (
                           Path.GetFileName(filePath),
                           $"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}{suffix}"),
                    .. from sheet in book.Sheets.Values
                       select (sheet.Name, $"{options.GeneratedName(sheet.Name)}Sheet"),
                    .. from table in book.Tables
                       let name = options.GeneratedName(table.Name)
                       from typeName in new[] { $"{name}Table", name }
                       select (table.Name, typeName)
                ]),
            []);

    /// <summary>
    /// シート・テーブルの生成プロパティ名が無効な場合、対応する元名を添えて報告します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> InvalidPropertyNameDiagnostics(
        IEnumerable<string> sourceNames,
        CodeGenerationOptions options) =>
        from sourceName in sourceNames
        let propertyName = options.GeneratedName(sourceName)
        where IsInvalidPropertyName(propertyName)
        select new CodeGenerationDiagnostic(true, propertyName, [sourceName], sourceName);

    static bool IsInvalidPropertyName(string propertyName) =>
        propertyName.IsEmpty()
            || propertyName.Any(it => " !\"#$%&'()*+,-./:;<=>?[\\]^`{|}~".Contains(it))
            || propertyName.Skip(1).Contains('@')
            || propertyName.Any(char.IsSurrogate)
            || propertyName == "@"
            || (propertyName.Any() && char.IsDigit(propertyName[0]))
            || (propertyName.Any() && char.GetUnicodeCategory(propertyName[0]) == UnicodeCategory.NonSpacingMark)
            || (propertyName.Length > 1 && propertyName[0] == '@' && char.IsDigit(propertyName[1]))
            || IsCSharpKeyword(propertyName);

    /// <summary>
    /// Book型の中で、定義名・ワークシート・Excelテーブルの生成プロパティ名同士、
    /// およびBook型名、生成・継承するAPIの予約名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> BookPropertyNameDiagnostics(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        NameCollisionDiagnostics(
            GroupMemberNames(
                Enumerable.Concat(BookDefinedNames(book, options), BookPropertyNames(book, options))),
            BookReservedNames(filePath));

    /// <summary>
    /// 同じSheetの定義名同士、および型名・継承した構造プロパティ名・ToStringとの衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> SheetDefinedNameDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from sheet in book.Sheets.Values
        from diagnostic in NameCollisionDiagnostics(
            GroupMemberNames(
                from definedName in book.DefinedNames
                where definedName.Worksheet?.Name == sheet.Name
                select (definedName.Name, options.SheetDefinedName(sheet, definedName))),
            SheetReservedNames(sheet, options))
        select diagnostic;

    /// <summary>
    /// 同じSheet型のローカル定義名とテーブルから生成するプロパティ名の衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> SheetDefinedNameTableDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from definedName in book.DefinedNames
        let sheet = definedName.Worksheet
        where sheet is not null
        from diagnostic in MemberNameCollisionDiagnostics(
            [(definedName.Name, options.SheetDefinedName(sheet, definedName))],
            from table in book.Tables
            where table.Worksheet.Name == sheet.Name
            select (table.Name, options.GeneratedName(table.Name)))
        select diagnostic;

    /// <summary>
    /// Sheetのテーブル由来プロパティと、所属するSheet型名・継承APIの予約名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> SheetTableNameDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from table in book.Tables
        from diagnostic in NameCollisionDiagnostics(
            [(options.GeneratedName(table.Name), new[] { table.Name })],
            SheetReservedNames(table.Worksheet, options))
        select diagnostic;

    /// <summary>
    /// Dataへ平坦化する定義名・テーブルのプロパティ同士、およびData型名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> DataPropertyNameDiagnostics(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        NameCollisionDiagnostics(
            GroupMemberNames(
                [
                    .. from definedName in book.DefinedNames
                       select (definedName.Name, options.GeneratedName(definedName)),
                    .. from table in book.Tables
                       select (table.Name, options.GeneratedName(table.Name))
                ]),
            [$"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Data"]);

    /// <summary>
    /// 不正な文字を含む列プロパティ名、指定の予約語と、列同士・行データ型名との名前衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> ColumnPropertyNameDiagnostics(
        Table table,
        CodeGenerationOptions options) =>
        [
            .. from column in table.Columns
               let propertyName = options.TableColumn(table, column)
               where propertyName.Any(it => it != '@' && char.GetUnicodeCategory(it) is not (
                       UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter
                       or UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter
                       or UnicodeCategory.OtherLetter or UnicodeCategory.LetterNumber
                       or UnicodeCategory.DecimalDigitNumber or UnicodeCategory.ConnectorPunctuation
                       or UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark
                       or UnicodeCategory.Format))
                   || propertyName.Skip(1).Any(it => it == '@')
                   || (propertyName.Any() && (char.IsDigit(propertyName[0])
                       || char.GetUnicodeCategory(propertyName[0]) is UnicodeCategory.NonSpacingMark
                           or UnicodeCategory.SpacingCombiningMark))
                   || IsCSharpKeyword(propertyName)
               select new CodeGenerationDiagnostic(true, propertyName, [column.Name], column.Name),
            .. NameCollisionDiagnostics(
                GroupMemberNames(
                    from column in table.Columns
                    select (column.Name, options.TableColumn(table, column))),
                [options.GeneratedName(table.Name)])
        ];

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

    /// <summary>
    /// Bookのシート・テーブル由来のプロパティを、元名と生成名の対応に揃えます。
    /// </summary>
    static IEnumerable<(string SourceName, string Name)> BookPropertyNames(
        Workbook book,
        CodeGenerationOptions options) =>
        from sourceName in
            Enumerable.Concat(
                from sheet in book.Sheets.Values select sheet.Name,
                from table in book.Tables select table.Name)
        select (sourceName, options.GeneratedName(sourceName));

    /// <summary>
    /// ブックスコープの定義名に、文脈付きの名前設定を適用します。
    /// </summary>
    static IEnumerable<(string SourceName, string Name)> BookDefinedNames(
        Workbook book,
        CodeGenerationOptions options) =>
        from definedName in book.DefinedNames
        where definedName.Worksheet is null
        select (definedName.Name, options.BookDefinedName(definedName));

    /// <summary>
    /// Bookの生成元の種類によらず、現在共通して予約している名前です。
    /// </summary>
    static string[] BookReservedNames(string filePath) =>
        [
            $"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Book",
            nameof(Workbook.Read), nameof(Workbook.Open), nameof(Workbook.Replace), "ValidateStructure",
            nameof(Workbook.Cell), nameof(Workbook.Range),
            nameof(Workbook.Tables), nameof(Workbook.DefinedNames), nameof(Workbook.Sheets),
            nameof(Workbook.Save), nameof(Workbook.SaveAs), nameof(Workbook.Close),
            nameof(Workbook.Dispose), nameof(Workbook.ReadTable)
        ];

    /// <summary>
    /// Sheetのプロパティの生成元によらず共通して予約する、所属型名と継承API名です。
    /// </summary>
    static string[] SheetReservedNames(Worksheet sheet, CodeGenerationOptions options) =>
        [
            $"{options.GeneratedName(sheet.Name)}Sheet", nameof(Worksheet.Cell), nameof(Worksheet.Range),
            nameof(Worksheet.Book), nameof(Worksheet.Name), nameof(Worksheet.Cells), nameof(Worksheet.ToString)
        ];

    /// <summary>
    /// 先頭のエスケープ表記と書式文字を除いて同じ識別子になるメンバーを、入力順を保ってまとめます。
    /// </summary>
    static IEnumerable<(string Name, string[] SourceNames)> GroupMemberNames(
        IEnumerable<(string SourceName, string Name)> members) =>
        from member in members
        group member.SourceName by member.Name.IdentifierValue into names
        select (names.Key, names.ToArray());

    /// <summary>
    /// 生成名と予約名を識別子名へ揃え、重複または予約名との衝突を一つの診断として返します。
    /// 定義名同士の重複を検証していない経路では、元名を一件ずつ渡します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> NameCollisionDiagnostics(
        IEnumerable<(string Name, string[] SourceNames)> names,
        IEnumerable<string> reservedNames) =>
        from name in names
        let identifier = name.Name.IdentifierValue
        where name.SourceNames.Length > 1
            || reservedNames.Select(it => it.IdentifierValue).Contains(identifier)
        select new CodeGenerationDiagnostic(true, identifier, name.SourceNames);

    /// <summary>
    /// 異なる生成元の間で同名になる組を、従来どおり元名の二者ごとに報告します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> MemberNameCollisionDiagnostics(
        IEnumerable<(string SourceName, string Name)> members,
        IEnumerable<(string SourceName, string Name)> otherMembers) =>
        from member in members
        from other in otherMembers
        where member.Name == other.Name
        select new CodeGenerationDiagnostic(true, member.Name, [member.SourceName, other.SourceName]);
}
