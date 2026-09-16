using System.Globalization;
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
        using System.Linq;
        using Marimo.SpreadSheetAsData;

        namespace {{options.Namespace}};

        {{BookDeclaration(filePath, book, options)}}
        {{BookDataDeclaration(filePath, book, options)}}
        {{ForEach(
            from sheet in book.Sheets.Values
            select SheetDeclaration(sheet, options))}}
        {{ForEach(
            from table in book.Tables
            select TableDeclaration(table, options))}}
        {{ForEach(
            from table in book.Tables
            select RowDeclaration(table, options))}}
        """;

    /// <summary>
    /// Excelブックを型付きで表すクラスの宣言を生成します。
    /// </summary>
    internal static string BookDeclaration(
        string filePath,
        Workbook book,
        CodeGenerationOptions options)
    {
        var bookFileName = Path.GetFileNameWithoutExtension(filePath);
        var bookFileIdentifier = bookFileName.ToCSharpIdentifier();

        return
            $$"""
        /// <summary>
        /// Excelブック「{{bookFileName}}」を型付きで表します。
        /// </summary>
        public partial class {{bookFileIdentifier}}Book : {{ReferencedTypeName("Workbook", book, options)}}
        {
            public {{bookFileIdentifier}}Book() : this({{StringLiteral(filePath)}})
            {
            }

            public {{bookFileIdentifier}}Book(string filePath) : base(filePath)
            {
            }

            /// <summary>
            /// 生成元のシート、テーブルとその列、両スコープの定義名を確認して、Excelブックを開きます。
            /// </summary>
            /// <exception cref="{{ReferencedTypeName("System.IO.InvalidDataException", book, options)}}">生成元のシート、テーブルとその列、両スコープの定義名のいずれかが存在しないか、テーブルの所属シートが異なるか、単一セルの定義名が複数セルを参照しています。</exception>
            public static new {{bookFileIdentifier}}Book Open(string filePath) =>
                ValidateStructure(new(filePath));

            {{bookFileIdentifier}}Book({{ReferencedTypeName("System.IO.Stream", book, options)}} stream) : base(stream)
            {
            }

            /// <summary>
            /// 生成元のシート、テーブルとその列、両スコープの定義名を確認して、Stream上のExcelブックを開きます。
            /// </summary>
            /// <exception cref="{{ReferencedTypeName("System.IO.InvalidDataException", book, options)}}">生成元のシート、テーブルとその列、両スコープの定義名のいずれかが存在しないか、テーブルの所属シートが異なるか、単一セルの定義名が複数セルを参照しています。</exception>
            public static new {{bookFileIdentifier}}Book Open({{ReferencedTypeName("System.IO.Stream", book, options)}} stream) =>
                ValidateStructure(new(stream));

            /// <summary>
            /// 生成元に対応するシート、テーブルとその列、両スコープの定義名を確認し、不一致時は開いたブックを破棄します。
            /// </summary>
            static {{bookFileIdentifier}}Book ValidateStructure({{bookFileIdentifier}}Book book)
            {
                if (new string[] { {{string.Join(
                    ", ",
                    from sheet in book.Sheets.Values
                    select StringLiteral(sheet.Name))}} }
                    .Any(sheetName => !book.Sheets.ContainsKey(sheetName))
                    || new (string SheetName, string TableName)[] { {{string.Join(
                        ", ",
                        from table in book.Tables
                        select $"({StringLiteral(table.Worksheet.Name)}, {StringLiteral(table.Name)})")}} }
                        .Any(name => !book.Tables.Any(
                            table => table.Name == name.TableName && table.Worksheet.Name == name.SheetName))
                    || new (string TableName, string ColumnName)[] { {{string.Join(
                        ", ",
                        from table in book.Tables
                        from column in table.Columns
                        select $"({StringLiteral(table.Name)}, {StringLiteral(column.Name)})")}} }
                        .Any(column => !book.Tables[column.TableName].Columns.Contains(column.ColumnName))
                    || new (string Name, bool RequiresSingleCell)[] { {{string.Join(
                        ", ",
                        from definedName in BookScopedDefinedNames(book)
                        select $"({StringLiteral(definedName.Name)}, {(IsSingleCellDefinedName(definedName) ? "true" : "false")})")}} }
                        .Any(name => !book.DefinedNames.Any(
                            definedName => definedName.Worksheet is null && definedName.Name == name.Name
                                && (!name.RequiresSingleCell
                                    || definedName.Range.TopLeftCell == definedName.Range.BottomRightCell)))
                    || new (string SheetName, string Name, bool RequiresSingleCell)[] { {{string.Join(
                        ", ",
                        from sheet in book.Sheets.Values
                        from definedName in SheetScopedDefinedNames(sheet)
                        select $"({StringLiteral(sheet.Name)}, {StringLiteral(definedName.Name)}, {(IsSingleCellDefinedName(definedName) ? "true" : "false")})")}} }
                        .Any(name => !book.DefinedNames.Any(
                            definedName => definedName.Worksheet?.Name == name.SheetName && definedName.Name == name.Name
                                && (!name.RequiresSingleCell
                                    || definedName.Range.TopLeftCell == definedName.Range.BottomRightCell))))
                {
                    book.Dispose();
                    throw new {{ReferencedTypeName("System.IO.InvalidDataException", book, options)}}();
                }

                return book;
            }

            /// <summary>
            /// Excelブック全体のデータを読み込みます。
            /// </summary>
            public {{bookFileIdentifier}}Data Read() =>
                base.Read<{{bookFileIdentifier}}Data>();

            /// <summary>
            /// Excelブック全体のデータを置換します。
            /// </summary>
            public void Replace({{bookFileIdentifier}}Data data) =>
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
        """;
    }

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
            .. from definedName in BookDataDefinedNames(book)
               select BookDataDefinedNamePropertyDeclaration(definedName, options),
            .. from table in book.Tables
               select BookDataTablePropertyDeclaration(table, options)
        ])}}
        }
        """;

    /// <summary>
    /// BookData型へ平坦化する定義名を、スコープごとに単一セル、セル範囲の順で列挙します。
    /// </summary>
    static IEnumerable<DefinedName> BookDataDefinedNames(Workbook book) =>
    [
        .. from definedName in BookScopedDefinedNames(book)
           where IsSingleCellDefinedName(definedName)
           select definedName,
        .. from definedName in BookScopedDefinedNames(book)
           where !IsSingleCellDefinedName(definedName)
           select definedName,
        .. from sheet in book.Sheets.Values
           from definedName in SheetScopedDefinedNames(sheet)
           where IsSingleCellDefinedName(definedName)
           select definedName,
        .. from sheet in book.Sheets.Values
           from definedName in SheetScopedDefinedNames(sheet)
           where !IsSingleCellDefinedName(definedName)
           select definedName
    ];

    /// <summary>
    /// ブックデータ型に、Excelテーブルの行データを表すプロパティを生成します。
    /// </summary>
    internal static string BookDataTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options)
    {
        var propertyName = options.GeneratedName(table.Name);
        var attributeDeclaration = table.Name.ToCSharpIdentifier() == propertyName
            ? ""
            : $"    [{ReferencedAttributeName(table.Worksheet.Book, options)}({StringLiteral(table.Name)})]{Environment.NewLine}";

        return $$"""

            /// <summary>
            /// Excelテーブル「{{table.Name}}」の行データを取得または設定します。
            /// </summary>
        {{attributeDeclaration}}    public IEnumerable<{{propertyName}}> {{propertyName}} { get; set; }
        """;
    }

    /// <summary>
    /// ブックデータ型に、定義名が表す値のプロパティを生成します。
    /// </summary>
    internal static string BookDataDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options)
    {
        var isSingleCell = IsSingleCellDefinedName(definedName);
        var propertyTypeName = isSingleCell
            ? CellValueTypeName(definedName.Range.TopLeftCell.Value)
            : "IEnumerable<IEnumerable<object?>>";
        var worksheet = definedName.Worksheet;
        var propertyName = worksheet is null
            ? options.BookDefinedName(definedName)
            : options.SheetDefinedName(worksheet, definedName);
        var attributeDeclaration = BookDataDefinedNameAttribute(
            definedName,
            propertyName,
            options);
        var scopeDescription = worksheet is null
            ? ""
            : $"ワークシート「{worksheet.Name}」の";
        var rangeDescription = isSingleCell ? "セル" : "セル範囲";
        var initializer = isSingleCell
            ? PropertyInitializer(propertyTypeName)
            : "";

        return
            $$"""

                /// <summary>
                /// {{scopeDescription}}定義名「{{definedName.Name}}」が表す{{rangeDescription}}の値を取得または設定します。
                /// </summary>
                {{attributeDeclaration}}public {{propertyTypeName}} {{propertyName}} { get; set; }{{initializer}}
            """;
    }

    /// <summary>
    /// 自動名前変換では対応できない定義名の属性を生成します。
    /// </summary>
    static string BookDataDefinedNameAttribute(
        DefinedName definedName,
        string propertyName,
        CodeGenerationOptions options)
    {
        // コンパイル後の識別子では書式文字が無視されるため、属性で元の定義名を保持します。
        if (propertyName == definedName.Name.ToCSharpIdentifier()
            && !propertyName.Any(it => char.GetUnicodeCategory(it) == UnicodeCategory.Format))
        {
            return "";
        }

        var worksheetArgument = definedName.Worksheet is null
            ? ""
            : $", WorksheetName = {StringLiteral(definedName.Worksheet.Name)}";

        return $"[{ReferencedAttributeName(definedName.Range.TopLeftCell.Book, options)}({StringLiteral(definedName.Name)}{worksheetArgument})]{Environment.NewLine}    ";
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
        public partial class {{options.GeneratedName(sheet.Name)}}Sheet : {{ReferencedTypeName("Worksheet", sheet.Book, options)}}
        {
            public {{options.GeneratedName(sheet.Name)}}Sheet({{ReferencedTypeName("Workbook", sheet.Book, options)}} book) : base(book, {{StringLiteral(sheet.Name)}})
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
    internal static string TableDeclaration(
        Table table,
        CodeGenerationOptions options)
    {
        var tableIdentifier = options.GeneratedName(table.Name);

        return $$"""

        /// <summary>
        /// Excelテーブル「{{table.Name}}」を型付きで表します。
        /// </summary>
        public partial class {{tableIdentifier}}Table : Table<{{tableIdentifier}}>
        {
            public {{tableIdentifier}}Table({{ReferencedTypeName("Table", table.Worksheet.Book, options)}} source) : base(source)
            {
            }
        }
        """;
    }

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
        public partial class {{options.GeneratedName(table.Name)}}
        {
        {{ForEach(
            from column in table.Columns
            where table.Rows.Any()
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
            {{ColumnAttributeDeclaration(column, propertyName, options)}}    public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// 生成プロパティ名とExcel列名が一致しない場合に、既存の型付きTableマッピングへ列名を伝える属性を生成します。
    /// </summary>
    internal static string ColumnAttributeDeclaration(
        TableColumn column,
        string propertyName,
        CodeGenerationOptions options) =>
        column.Name == propertyName
            ? ""
            : $"    [{ReferencedAttributeName(column.Table.Worksheet.Book, options)}({StringLiteral(column.Name)})]{Environment.NewLine}";

    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    internal static string ColumnPropertyTypeName(Table table, TableColumn column)
    {
        var values = (
            from row in table.Rows
            select row[column].Value
        ).ToArray();

        if (values.All(it => it is string or BlankValue))
        {
            return "string";
        }

        if (values.All(CanConvertToInt32))
        {
            return "int";
        }

        if (values.All(it => it is BlankValue || CanConvertToInt32(it)))
        {
            return "int?";
        }

        if (values.All(it => it is double))
        {
            return "double";
        }

        if (values.All(it => it is double or BlankValue))
        {
            return "double?";
        }

        if (values.All(it => it is bool))
        {
            return "bool";
        }

        if (values.All(it => it is bool or BlankValue))
        {
            return "bool?";
        }

        return "dynamic";

        // セル値がintの範囲に収まる整数値かどうかを判定します。
        static bool CanConvertToInt32(object value) =>
            value is double number
            && double.IsInteger(number)
            && number is >= int.MinValue and <= int.MaxValue;
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
            bool => "bool",
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
        CellDefinedNamePropertyDeclaration(definedName, options.BookDefinedName(definedName), "");

    /// <summary>
    /// スコープごとに決めた名前と説明を使い、単一セル値への読み書きを生成します。
    /// </summary>
    static string CellDefinedNamePropertyDeclaration(
        DefinedName definedName,
        string propertyName,
        string scopeDescription) =>
        $$"""

            /// <summary>
            /// {{scopeDescription}}定義名「{{definedName.Name}}」が表すセルの値を取得または設定します。
            /// </summary>
            public dynamic {{propertyName}}
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
        CellRangeDefinedNamePropertyDeclaration(definedName, options.BookDefinedName(definedName), "");

    /// <summary>
    /// スコープごとに決めた名前と説明を使い、セル範囲の値への読み書きを生成します。
    /// </summary>
    static string CellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        string propertyName,
        string scopeDescription) =>
        $$"""

            /// <summary>
            /// {{scopeDescription}}定義名「{{definedName.Name}}」が表すセル範囲の値を取得または設定します。
            /// </summary>
            public IEnumerable<IEnumerable<object?>> {{propertyName}}
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
        CellDefinedNamePropertyDeclaration(
            definedName,
            options.SheetDefinedName(sheet, definedName),
            $"ワークシート「{sheet.Name}」の");

    /// <summary>
    /// ワークシートスコープのセル範囲定義名が表す値を読み書きするプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetCellRangeDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        CellRangeDefinedNamePropertyDeclaration(
            definedName,
            options.SheetDefinedName(sheet, definedName),
            $"ワークシート「{sheet.Name}」の");

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
    /// 行データ型が参照名の先頭を隠す場合だけ、ルートからの完全修飾名を返します。
    /// その他の生成型には接尾辞が付くため、ここで扱う既存型名とは衝突しません。
    /// </summary>
    static string ReferencedTypeName(string typeName, Workbook book, CodeGenerationOptions options) =>
        !book.Tables.Any(table => options.GeneratedName(table.Name).IdentifierValue == typeName.Split('.')[0])
            ? typeName
        : typeName.Contains('.')
            ? $"global::{typeName}"
        : $"global::Marimo.SpreadSheetAsData.{typeName}";

    /// <summary>
    /// 行データ型が属性クラス名を隠す場合だけ完全修飾し、それ以外は属性の短縮表記を使います。
    /// </summary>
    static string ReferencedAttributeName(Workbook book, CodeGenerationOptions options)
    {
        var typeName = ReferencedTypeName(nameof(SpreadSheetNameAttribute), book, options);
        return typeName == nameof(SpreadSheetNameAttribute) ? "SpreadSheetName" : typeName;
    }

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
