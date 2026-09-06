using Marimo.SpreadSheetAsData;
namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// Workbookラッパー生成で使用するテンプレート部品です。
/// </summary>
static class WorkbookWrapperComponents
{
    /// <summary>
    /// 生成するC#ソースファイル全体を表すテンプレート部品です。
    /// </summary>
    internal static string SourceFile(
        string filePath,
        CodeGenerationOptions options,
        Workbook book) =>
        $$"""
        using System.Collections.Generic;
        using Marimo.SpreadSheetAsData;

        namespace {{options.Namespace}};

        /// <summary>
        /// Excelブック「{{Path.GetFileNameWithoutExtension(filePath)}}」を型付きで表します。
        /// </summary>
        public partial class {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book : Workbook
        {
            public {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book() : this({{StringLiteral(filePath)}})
            {
            }

            public {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book(string filePath) : base(filePath)
            {
            }

            public static new {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book Open(string filePath) =>
                new(filePath);

            {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book(System.IO.Stream stream) : base(stream)
            {
            }

            public static new {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Book Open(System.IO.Stream stream) =>
                new(stream);

            /// <summary>
            /// Excelブック全体のデータを読み込みます。
            /// </summary>
            public new {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Data Read() =>
                base.Read<{{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Data>();

            /// <summary>
            /// Excelブック全体のデータを置換します。
            /// </summary>
            public new void Replace({{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Data data) =>
                base.Replace(data);
        {{ForEach([
            .. from definedName in BookScopedDefinedNames(book)
               where IsSingleCellDefinedName(definedName)
               select BookCellDefinedNamePropertyDeclaration(definedName, options),
            .. from definedName in BookScopedDefinedNames(book)
               where !IsSingleCellDefinedName(definedName)
               select BookCellRangeDefinedNamePropertyDeclaration(definedName, options),
            .. from sheet in book.Sheets.Values
               select SheetPropertyDeclaration(sheet, options),
            .. from table in book.Tables
               select BookTablePropertyDeclaration(table, options)
        ])}}
        }
        {{BookDataDeclaration(filePath, book, options)}}
        {{ForEach(
            from sheet in book.Sheets.Values
            select SheetDeclaration(sheet, options))}}
        {{ForEach(
            from table in book.Tables
            select TableDeclaration(table))}}
        {{ForEach(
            from table in book.Tables
            select RowDeclaration(table, options))}}
        """;

    /// <summary>
    /// Excelブック全体のデータを表す型の宣言を生成します。
    /// </summary>
    internal static string BookDataDeclaration(
        string filePath,
        Workbook book,
        CodeGenerationOptions options) =>
        $$"""

        /// <summary>
        /// Excelブック「{{Path.GetFileNameWithoutExtension(filePath)}}」のデータを表します。
        /// </summary>
        public partial class {{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}}Data
        {
        {{ForEach([
            .. from definedName in BookScopedDefinedNames(book)
               where IsSingleCellDefinedName(definedName)
               select BookDataCellPropertyDeclaration(definedName, options),
            .. from definedName in BookScopedDefinedNames(book)
               where !IsSingleCellDefinedName(definedName)
               select BookDataCellRangePropertyDeclaration(definedName, options),
            .. from sheet in book.Sheets.Values
               from definedName in SheetScopedDefinedNames(sheet)
               where IsSingleCellDefinedName(definedName)
               select BookDataSheetCellPropertyDeclaration(sheet, definedName, options),
            .. from sheet in book.Sheets.Values
               from definedName in SheetScopedDefinedNames(sheet)
               where !IsSingleCellDefinedName(definedName)
               select BookDataSheetCellRangePropertyDeclaration(sheet, definedName, options)
        ])}}
        }
        """;

    /// <summary>
    /// ブックデータ型に、ブックスコープの単一セル定義名が表すプロパティを生成します。
    /// </summary>
    internal static string BookDataCellPropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options)
    {
        var propertyTypeName = CellValueTypeName(definedName.Range.TopLeftCell.Value);
        var propertyName = options.BookDefinedName(definedName);

        return
            $$"""

                /// <summary>
                /// 定義名「{{definedName.Name}}」が表すセルの値を取得または設定します。
                /// </summary>
                public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// ブックデータ型に、ブックスコープの複数セル定義名が表すプロパティを生成します。
    /// </summary>
    internal static string BookDataCellRangePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options)
    {
        var propertyName = options.BookDefinedName(definedName);

        return
            $$"""

                /// <summary>
                /// 定義名「{{definedName.Name}}」が表すセル範囲の値を取得または設定します。
                /// </summary>
                public IEnumerable<IEnumerable<object?>> {{propertyName}} { get; set; }
            """;
    }

    /// <summary>
    /// ブックデータ型に、シートローカルの単一セル定義名が表すプロパティを生成します。
    /// </summary>
    internal static string BookDataSheetCellPropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options)
    {
        var propertyTypeName = CellValueTypeName(definedName.Range.TopLeftCell.Value);
        var propertyName = options.SheetDefinedName(sheet, definedName);

        return
            $$"""

                /// <summary>
                /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセルの値を取得または設定します。
                /// </summary>
                public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// ブックデータ型に、シートローカルの複数セル定義名が表すプロパティを生成します。
    /// </summary>
    internal static string BookDataSheetCellRangePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options)
    {
        var propertyName = options.SheetDefinedName(sheet, definedName);

        return
            $$"""

                /// <summary>
                /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセル範囲の値を取得します。
                /// </summary>
                public IEnumerable<IEnumerable<object?>> {{propertyName}} { get; }
            """;
    }

    /// <summary>
    /// Book型から指定ワークシート型を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetPropertyDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(sheet.Name)}}Sheet {{options.GeneratedName(sheet.Name)}} => new(this);
        """;

    /// <summary>
    /// Book型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string BookTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// Excelテーブル「{{table.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(table.Name)}}Table {{options.GeneratedName(table.Name)}} => new(Tables[{{StringLiteral(table.Name)}}]);
        """;

    /// <summary>
    /// ワークシートを表す派生Sheet型の宣言を生成します。
    /// </summary>
    internal static string SheetDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options)
        => $$"""

        /// <summary>
        /// ワークシート「{{sheet.Name}}」を型付きで表します。
        /// </summary>
        public partial class {{options.GeneratedName(sheet.Name)}}Sheet : Worksheet
        {
            public {{options.GeneratedName(sheet.Name)}}Sheet(Workbook book) : base(book, {{StringLiteral(sheet.Name)}})
            {
            }
        {{ForEach([
            .. from definedName in SheetScopedDefinedNames(sheet)
               where IsSingleCellDefinedName(definedName)
               select SheetCellDefinedNamePropertyDeclaration(sheet, definedName, options),
            .. from definedName in SheetScopedDefinedNames(sheet)
               where !IsSingleCellDefinedName(definedName)
               select SheetCellRangeDefinedNamePropertyDeclaration(sheet, definedName, options),
            .. from table in sheet.Book.Tables
               where table.Worksheet.Name == sheet.Name
               select SheetTablePropertyDeclaration(table, options)
        ])}}
        }
        """;

    /// <summary>
    /// Excelテーブルを表す派生Table型の宣言を生成します。
    /// </summary>
    internal static string TableDeclaration(Table table)
        => $$"""

        /// <summary>
        /// Excelテーブル「{{table.Name}}」を型付きで表します。
        /// </summary>
        public partial class {{table.Name.ToCSharpIdentifier()}}Table : Table<{{table.Name.ToCSharpIdentifier()}}>
        {
            public {{table.Name.ToCSharpIdentifier()}}Table(Table source) : base(source)
            {
            }
        }
        """;

    /// <summary>
    /// Excelテーブルの1行を表す行データ型の宣言を生成します。
    /// </summary>
    internal static string RowDeclaration(
        Table table,
        CodeGenerationOptions options)
        => $$"""

        /// <summary>
        /// Excelテーブル「{{table.Name}}」の1行を表します。
        /// </summary>
        public partial class {{table.Name.ToCSharpIdentifier()}}
        {
        {{ForEach(
            from column in table.Columns
            select RowPropertyDeclaration(table, column, options))}}
        }
        """;

    /// <summary>
    /// Excelテーブル列に対応する行データプロパティ宣言を生成します。
    /// </summary>
    internal static string RowPropertyDeclaration(
        Table table,
        TableColumn column,
        CodeGenerationOptions options)
    {
        var propertyTypeName = ColumnPropertyTypeName(table, column);
        var propertyName = options.TableColumn(table, column);

        return
            $$"""

                /// <summary>
                /// Excel列「{{column.Name}}」の値を取得または設定します。
                /// </summary>
            {{ColumnAttributeDeclaration(column, propertyName)}}    public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// 生成プロパティ名とExcel列名が一致しない場合に、既存の型付きTableマッピングへ列名を伝える属性を生成します。
    /// </summary>
    internal static string ColumnAttributeDeclaration(TableColumn column, string propertyName) =>
        column.Name == propertyName
            ? ""
            : $"    [SpreadsheetColumn({StringLiteral(column.Name)})]{Environment.NewLine}";

    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    internal static string ColumnPropertyTypeName(Table table, TableColumn column)
    {
        var values = (
            from row in table.Rows
            select row[column].Value
        ).ToArray();

        if (values.All(it => it is string))
        {
            return "string";
        }

        if (values.All(it => it is double number && double.IsInteger(number)))
        {
            return "int";
        }

        if (values.All(it => it is double))
        {
            return "double";
        }

        return "object?";
    }

    /// <summary>
    /// 生成プロパティがコンパイル時の初期化警告を出さないための初期値を返します。
    /// </summary>
    internal static string PropertyInitializer(string propertyTypeName) =>
        propertyTypeName == "string"
            ? " = \"\";"
            : "";

    /// <summary>
    /// セルの現在値を表すC#プロパティ型名を返します。
    /// </summary>
    internal static string CellValueTypeName(object value) =>
        value switch
        {
            string => "string",
            double => "double",
            _ => "dynamic"
        };

    /// <summary>
    /// ブック全体から参照できる定義名だけを選びます。
    /// </summary>
    static IEnumerable<DefinedName> BookScopedDefinedNames(Workbook book) =>
        from definedName in book.DefinedNames
        where definedName.Worksheet == null
        select definedName;

    /// <summary>
    /// 指定ワークシートだけで参照できる定義名を選びます。
    /// </summary>
    static IEnumerable<DefinedName> SheetScopedDefinedNames(Worksheet sheet) =>
        from definedName in sheet.Book.DefinedNames
        where definedName.Worksheet?.Name == sheet.Name
        select definedName;

    /// <summary>
    /// 定義名が単一セルを指すかどうかを判定します。
    /// </summary>
    static bool IsSingleCellDefinedName(DefinedName definedName) =>
        definedName.Range.TopLeftCell == definedName.Range.BottomRightCell;

    /// <summary>
    /// ブックスコープの単一セル定義名が表す値を読み書きするプロパティ宣言を生成します。
    /// </summary>
    internal static string BookCellDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// 定義名「{{definedName.Name}}」が表すセルの値を取得または設定します。
            /// </summary>
            public dynamic {{options.BookDefinedName(definedName)}}
            {
                get => Cell[{{StringLiteral(definedName.Name)}}].Value;
                set => Cell[{{StringLiteral(definedName.Name)}}].Value = value;
            }
        """;

    /// <summary>
    /// ブックスコープのセル範囲定義名が表す値を読み書きするプロパティ宣言を生成します。
    /// </summary>
    internal static string BookCellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// 定義名「{{definedName.Name}}」が表すセル範囲の値を取得または設定します。
            /// </summary>
            public IEnumerable<IEnumerable<object?>> {{options.BookDefinedName(definedName)}}
            {
                get => Range[{{StringLiteral(definedName.Name)}}].Values;
                set => Range[{{StringLiteral(definedName.Name)}}].Values = value;
            }
        """;

    /// <summary>
    /// ワークシートスコープの単一セル定義名が表す値を読み書きするプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetCellDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセルの値を取得または設定します。
            /// </summary>
            public dynamic {{options.SheetDefinedName(sheet, definedName)}}
            {
                get => Cell[{{StringLiteral(definedName.Name)}}].Value;
                set => Cell[{{StringLiteral(definedName.Name)}}].Value = value;
            }
        """;

    /// <summary>
    /// ワークシートスコープのセル範囲定義名が表す値を読み書きするプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetCellRangeDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセル範囲の値を取得または設定します。
            /// </summary>
            public IEnumerable<IEnumerable<object?>> {{options.SheetDefinedName(sheet, definedName)}}
            {
                get => Range[{{StringLiteral(definedName.Name)}}].Values;
                set => Range[{{StringLiteral(definedName.Name)}}].Values = value;
            }
        """;

    /// <summary>
    /// Sheet型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// Excelテーブル「{{table.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(table.Name)}}Table {{options.GeneratedName(table.Name)}} => new(Book.Tables[{{StringLiteral(table.Name)}}]);
        """;

    /// <summary>
    /// 複数のテンプレート部品を、生成ソース上の行単位で連結します。
    /// </summary>
    internal static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    /// <summary>
    /// 生成コード内へ埋め込む文字列リテラルを作ります。
    /// </summary>
    internal static string StringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
