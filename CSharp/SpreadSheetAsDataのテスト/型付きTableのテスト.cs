using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class 型付きTableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";
    const string MappingTableName = "型付き行マッピング";
    const string PropertyNameMappingTableName = "プロパティ名マッピング";

    readonly Workbook book;
    readonly Table<TestMappedRow> tested;

    public 型付きTableのテスト()
    {
        book = Workbook.Open(TestFilePath);
        tested = book.ReadTable<TestMappedRow>(MappingTableName);
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void 型付きTableは属性で指定した列をプロパティへ設定して各データ行を順に列挙します()
    {
        tested.Should().BeEquivalentTo(
            [
                new TestMappedRow
                {
                    IntegerValue = 1,
                    FloatingPointValue = 4.4,
                    TextValue = "さしすせそ"
                },
                new TestMappedRow
                {
                    IntegerValue = 2,
                    FloatingPointValue = 5.5,
                    TextValue = "たちつてと"
                },
                new TestMappedRow
                {
                    IntegerValue = 3,
                    FloatingPointValue = 6.6,
                    TextValue = "なにぬねの"
                }
            ],
            options => options.WithStrictOrdering());
    }

    [Fact]
    public void 型付きTableはデータ行がない場合に空の列挙になります()
    {
        book.ReadTable<EmptyTableRow>("空行マッピング")
            .Should().BeEmpty();
    }

    [Fact]
    public void 型付きTableは属性がないプロパティ名を列名として使用します()
    {
        book.ReadTable<PropertyNameMappedRow>(PropertyNameMappingTableName)
            .Should().BeEquivalentTo(
                [
                    new PropertyNameMappedRow
                    {
                        @float = 1.1,
                        @string = "あいうえお"
                    },
                    new PropertyNameMappedRow
                    {
                        @float = 2.2,
                        @string = "かきくけこ"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void 型付きTableはマッピング先に対応プロパティがない列を無視します()
    {
        book.ReadTable<IntegerOnlyRow>(MappingTableName)
            .Select(row => row.IntegerValue)
            .Should().Equal(1, 2, 3);
    }

    [Fact]
    public void 型付きTableは属性で指定した列が存在しない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<MissingColumnRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(typeof(MissingColumnRow));
        thrown.Which.ColumnName.Should().Be("missing");
        thrown.Which.PropertyName.Should().Be(nameof(MissingColumnRow.Value));
        thrown.Which.PropertyType.Should().Be(typeof(int));
        thrown.Which.WorksheetRowIndex.Should().BeNull();
        thrown.Which.SourceValue.Should().BeNull();
    }

    [Fact]
    public void 型付きTableはセル値をプロパティ型へ変換できない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<StringAsIntegerRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(typeof(StringAsIntegerRow));
        thrown.Which.ColumnName.Should().Be("文字列");
        thrown.Which.PropertyName.Should().Be(nameof(StringAsIntegerRow.Value));
        thrown.Which.PropertyType.Should().Be(typeof(int));
        thrown.Which.WorksheetRowIndex.Should().Be(7U);
        thrown.Which.SourceValue.Should().Be("さしすせそ");
    }

    [Fact]
    public void 型付きTableは小数をintへ変換しようとした場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<FloatingPointAsIntegerRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(
            typeof(FloatingPointAsIntegerRow));
        thrown.Which.ColumnName.Should().Be("数値");
        thrown.Which.PropertyName.Should().Be(
            nameof(FloatingPointAsIntegerRow.Value));
        thrown.Which.PropertyType.Should().Be(typeof(int));
        thrown.Which.WorksheetRowIndex.Should().Be(7U);
        thrown.Which.SourceValue.Should().Be(4.4);
    }

    [Fact]
    public void 型付きTableは複数のプロパティが同じ列を指定した場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<DuplicateColumnRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(typeof(DuplicateColumnRow));
        thrown.Which.ColumnName.Should().Be("数値2");
        thrown.Which.WorksheetRowIndex.Should().BeNull();
        thrown.Which.SourceValue.Should().BeNull();
    }

    [Fact]
    public void 型付きTableは属性を付けたプロパティにpublicなsetterがない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<AttributedPropertyWithoutPublicSetterRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(
            typeof(AttributedPropertyWithoutPublicSetterRow));
        thrown.Which.ColumnName.Should().Be("数値2");
        thrown.Which.PropertyName.Should().Be(
            nameof(
                AttributedPropertyWithoutPublicSetterRow.IntegerValue));
        thrown.Which.PropertyType.Should().Be(typeof(int));
        thrown.Which.WorksheetRowIndex.Should().BeNull();
        thrown.Which.SourceValue.Should().BeNull();
    }

    [Fact]
    public void 型付きTableは属性のない読み取り専用プロパティを無視します()
    {
        var rows = book
            .ReadTable<RowWithReadOnlyProperty>(MappingTableName)
            .ToArray();

        rows
            .Select(row => row.IntegerValue)
            .Should().Equal(1, 2, 3);

        rows
            .Select(row => row.Description)
            .Should().OnlyContain(value => value == "computed");
    }

    [Fact]
    public void 型付きTableはpublicな引数なしコンストラクターがない型では失敗します()
    {
        var action = () =>
        {
            book.ReadTable<RowWithoutPublicParameterlessConstructor>(MappingTableName)
                .ToArray();
        };

        var thrown = action
            .Should().Throw<TableMappingException>();

        thrown.Which.TableName.Should().Be(MappingTableName);
        thrown.Which.MappingType.Should().Be(
            typeof(RowWithoutPublicParameterlessConstructor));
        thrown.Which.ColumnName.Should().BeNull();
        thrown.Which.PropertyName.Should().BeNull();
        thrown.Which.PropertyType.Should().BeNull();
        thrown.Which.WorksheetRowIndex.Should().BeNull();
        thrown.Which.SourceValue.Should().BeNull();
    }

    public sealed class PropertyNameMappedRow
    {
        public double @float { get; set; }

        public string @string { get; set; } = "";
    }

    public sealed class IntegerOnlyRow
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; set; }
    }

    public sealed class MissingColumnRow
    {
        [SpreadsheetColumn("missing")]
        public int Value { get; set; }
    }

    public sealed class StringAsIntegerRow
    {
        [SpreadsheetColumn("文字列")]
        public int Value { get; set; }
    }

    public sealed class FloatingPointAsIntegerRow
    {
        [SpreadsheetColumn("数値")]
        public int Value { get; set; }
    }

    public sealed class EmptyTableRow
    {
        [SpreadsheetColumn("列1")]
        public string FirstValue { get; set; } = "";

        [SpreadsheetColumn("列2")]
        public string SecondValue { get; set; } = "";
    }

    public sealed class DuplicateColumnRow
    {
        [SpreadsheetColumn("数値2")]
        public int FirstValue { get; set; }

        [SpreadsheetColumn("数値2")]
        public int SecondValue { get; set; }
    }

    public sealed class AttributedPropertyWithoutPublicSetterRow
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; private set; }
    }

    public sealed class RowWithReadOnlyProperty
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; set; }

        public string Description => "computed";
    }

    public sealed class RowWithoutPublicParameterlessConstructor
    {
        public RowWithoutPublicParameterlessConstructor(int value)
        {
            Value = value;
        }

        [SpreadsheetColumn("数値2")]
        public int Value { get; set; }
    }
}
