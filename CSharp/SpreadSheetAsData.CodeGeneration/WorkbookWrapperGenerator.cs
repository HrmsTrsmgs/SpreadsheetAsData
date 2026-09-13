using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.WorkbookWrapperComponents;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// Excelブックから、SpreadsheetAsDataの型付きラッパーコードを生成します。
/// </summary>
public static class WorkbookWrapperGenerator
{
    /// <summary>
    /// 指定したExcelブックからC#ソースコードを生成します。
    /// </summary>
    /// <param name="filePath">生成元のExcelブックのパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    public static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null)
    {
        var options = new CodeGenerationOptions();
        configure?.Invoke(options);

        using var book = Workbook.Open(filePath);
        return
        [
            SourceFile(filePath, options, book)
        ];
    }

    /// <summary>
    /// 指定したExcelブックを解析し、コード生成前に検出できる問題を診断します。
    /// </summary>
    /// <param name="filePath">診断対象のExcelブックのパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>検出された診断情報。</returns>
    public static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null)
    {
        var options = new CodeGenerationOptions();
        configure?.Invoke(options);

        using var book = Workbook.Open(filePath);
        return
        [
            .. InvalidBookSheetPropertyNameDiagnostics(book, options),
            .. BookPropertyNameDiagnostics(filePath, book, options),
            .. BookDefinedNameDiagnostics(filePath, book, options),
            .. BookDefinedNameSheetDiagnostics(book, options),
            .. SheetDefinedNameDiagnostics(book, options),
            .. from table in book.Tables
               from diagnostic in ColumnPropertyNameDiagnostics(table, options)
               select diagnostic
        ];
    }

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
    /// Book型の中で、ワークシートとExcelテーブルの生成プロパティ名同士、
    /// およびBook型名とRead・Open・Replace・ValidateStructureメソッド名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> BookPropertyNameDiagnostics(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        from sourceNamesByPropertyName in
            from sourceName in
                Enumerable.Concat(
                    from sheet in book.Sheets.Values select sheet.Name,
                    from table in book.Tables select table.Name)
            group sourceName by options.GeneratedName(sourceName)
        where sourceNamesByPropertyName.Skip(1).Any()
            || sourceNamesByPropertyName.Key == $"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Book"
            || sourceNamesByPropertyName.Key is nameof(Workbook.Read) or nameof(Workbook.Open)
                or nameof(Workbook.Replace) or "ValidateStructure"
        select new CodeGenerationDiagnostic(
            true,
            sourceNamesByPropertyName.Key,
            [.. sourceNamesByPropertyName]);

    /// <summary>
    /// Book型名、自身が宣言するメソッド名、継承したCell・Range・Tables・DefinedNames
    /// プロパティ名との定義名の衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> BookDefinedNameDiagnostics(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        from definedName in book.DefinedNames
        where definedName.Worksheet is null
        let propertyName = options.BookDefinedName(definedName)
        where propertyName == $"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Book"
            || propertyName is nameof(Workbook.Read) or nameof(Workbook.Open)
                or nameof(Workbook.Replace) or "ValidateStructure"
                or nameof(Workbook.Cell) or nameof(Workbook.Range) or nameof(Workbook.Tables)
                or nameof(Workbook.DefinedNames)
        select new CodeGenerationDiagnostic(
            true,
            propertyName,
            [definedName.Name]);

    /// <summary>
    /// Book型の定義名とシートから生成するプロパティ名の衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> BookDefinedNameSheetDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from definedName in book.DefinedNames
        where definedName.Worksheet is null
        let propertyName = options.BookDefinedName(definedName)
        from sheet in book.Sheets.Values
        where propertyName == options.GeneratedName(sheet.Name)
        select new CodeGenerationDiagnostic(
            true,
            propertyName,
            [definedName.Name, sheet.Name]);

    /// <summary>
    /// Sheet型名、継承したCell・Rangeプロパティ名とのシートローカル定義名の衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> SheetDefinedNameDiagnostics(
        Workbook book,
        CodeGenerationOptions options) =>
        from definedName in book.DefinedNames
        let sheet = definedName.Worksheet
        where sheet is not null
        let propertyName = options.SheetDefinedName(sheet, definedName)
        where propertyName == $"{options.GeneratedName(sheet.Name)}Sheet"
            || propertyName is nameof(Worksheet.Cell) or nameof(Worksheet.Range)
        select new CodeGenerationDiagnostic(
            true,
            propertyName,
            [definedName.Name]);

    /// <summary>
    /// 列プロパティ同士の名前衝突と、行データ型名との衝突を検出します。
    /// </summary>
    static IEnumerable<CodeGenerationDiagnostic> ColumnPropertyNameDiagnostics(
        Table table,
        CodeGenerationOptions options) =>
        from columnsByPropertyName in
            from column in table.Columns
            group column.Name by options.TableColumn(table, column)
        where columnsByPropertyName.Skip(1).Any()
            || columnsByPropertyName.Key == options.GeneratedName(table.Name)
        select new CodeGenerationDiagnostic(
            true,
            columnsByPropertyName.Key,
            [.. columnsByPropertyName]);
}
