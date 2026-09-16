using Marimo.SpreadSheetAsData;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

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
            .. InvalidBookSheetPropertyNameDiagnostics(book, options),
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
    /// Book型のプロパティ名を生成できないワークシート名を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> InvalidBookSheetPropertyNameDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from sheet in book.Sheets.Values
        let propertyName = options.GeneratedName(sheet.Name)
        where propertyName.IsEmpty()
        select new CodeGenerationDiagnostic(
            true,
            propertyName,
            [sheet.Name],
            sheet.Name);

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
                       let sheet = definedName.Worksheet
                       let propertyName = sheet is null
                           ? options.BookDefinedName(definedName)
                           : options.SheetDefinedName(sheet, definedName)
                       select (definedName.Name, propertyName),
                    .. from table in book.Tables
                       select (table.Name, options.GeneratedName(table.Name))
                ]),
            [$"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Data"]);

    /// <summary>
    /// 列プロパティ同士の名前衝突と、行データ型名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> ColumnPropertyNameDiagnostics(
        Table table,
        CodeGenerationOptions options) =>
        NameCollisionDiagnostics(
            GroupMemberNames(
                from column in table.Columns
                select (column.Name, options.TableColumn(table, column))),
            [options.GeneratedName(table.Name)]);

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
