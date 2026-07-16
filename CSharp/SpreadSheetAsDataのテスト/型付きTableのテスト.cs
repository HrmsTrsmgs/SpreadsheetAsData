using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class 型付きTableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";
    const string MappingTableName = "テーブル6";

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
                    IntegerValue = 4,
                    FloatingPointValue = 4.4,
                    TextValue = "か"
                },
                new TestMappedRow
                {
                    IntegerValue = 2,
                    FloatingPointValue = 2.2,
                    TextValue = "き"
                },
                new TestMappedRow
                {
                    IntegerValue = 3,
                    FloatingPointValue = 3.3,
                    TextValue = "く"
                },
                new TestMappedRow
                {
                    IntegerValue = 10,
                    FloatingPointValue = 11,
                    TextValue = "け"
                },
                new TestMappedRow
                {
                    IntegerValue = 1,
                    FloatingPointValue = 1.1,
                    TextValue = "こ"
                },
                new TestMappedRow
                {
                    IntegerValue = 9,
                    FloatingPointValue = 9.9,
                    TextValue = "さ"
                },
                new TestMappedRow
                {
                    IntegerValue = 6,
                    FloatingPointValue = 6.6,
                    TextValue = "あ"
                },
                new TestMappedRow
                {
                    IntegerValue = 8,
                    FloatingPointValue = 8.8,
                    TextValue = "い"
                },
                new TestMappedRow
                {
                    IntegerValue = 5,
                    FloatingPointValue = 5.5,
                    TextValue = "う"
                },
                new TestMappedRow
                {
                    IntegerValue = 5,
                    FloatingPointValue = 6.6,
                    TextValue = "え"
                },
                new TestMappedRow
                {
                    IntegerValue = 7,
                    FloatingPointValue = 7.7,
                    TextValue = "お"
                }
            ],
            options => options.WithStrictOrdering());
    }

    [Fact]
    public void 型付きTableはデータ行がない場合に空の列挙になります()
    {
        book.ReadTable<TestMappedRow>("テーブル3")
            .Should()
            .BeEmpty();
    }

    [Fact]
    public void 型付きTableは属性がないプロパティ名を列名として使用します()
    {
        book.ReadTable<PropertyNameMappedRow>(MappingTableName)
            .Should()
            .BeEquivalentTo(
                [
                    new PropertyNameMappedRow
                    {
                        @int = 4,
                        @float = 4.4,
                        @string = "か"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 2,
                        @float = 2.2,
                        @string = "き"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 3,
                        @float = 3.3,
                        @string = "く"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 10,
                        @float = 11,
                        @string = "け"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 1,
                        @float = 1.1,
                        @string = "こ"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 9,
                        @float = 9.9,
                        @string = "さ"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 6,
                        @float = 6.6,
                        @string = "あ"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 8,
                        @float = 8.8,
                        @string = "い"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 5,
                        @float = 5.5,
                        @string = "う"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 5,
                        @float = 6.6,
                        @string = "え"
                    },
                    new PropertyNameMappedRow
                    {
                        @int = 7,
                        @float = 7.7,
                        @string = "お"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void 型付きTableはマッピング先に対応プロパティがない列を無視します()
    {
        book.ReadTable<IntegerOnlyRow>(MappingTableName)
            .Select(row => row.IntegerValue)
            .Should()
            .Equal(4, 2, 3, 10, 1, 9, 6, 8, 5, 5, 7);
    }

    [Fact]
    public void 型付きTableは属性で指定した列が存在しない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<MissingColumnRow>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(typeof(MissingColumnRow));
        exception.ColumnName.Should().Be("missing");
        exception.PropertyName.Should().Be(nameof(MissingColumnRow.Value));
        exception.PropertyType.Should().Be(typeof(int));
        exception.WorksheetRowIndex.Should().BeNull();
        exception.SourceValue.Should().BeNull();
    }

    [Fact(Skip = "セル値を対象プロパティ型へ変換できない場合のマッピングエラーを実装するときに解除する。")]
    public void 型付きTableはセル値をプロパティ型へ変換できない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<StringAsIntegerRow>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(typeof(StringAsIntegerRow));
        exception.ColumnName.Should().Be("string");
        exception.PropertyName.Should().Be(nameof(StringAsIntegerRow.Value));
        exception.PropertyType.Should().Be(typeof(int));
        exception.WorksheetRowIndex.Should().Be(2U);
        exception.SourceValue.Should().Be("か");
    }

    [Fact(Skip = "小数をintへ暗黙に丸めず変換失敗とする規則を実装するときに解除する。")]
    public void 型付きTableは小数をintへ変換しようとした場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<FloatingPointAsIntegerRow>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(
            typeof(FloatingPointAsIntegerRow));
        exception.ColumnName.Should().Be("float");
        exception.PropertyName.Should().Be(
            nameof(FloatingPointAsIntegerRow.Value));
        exception.PropertyType.Should().Be(typeof(int));
        exception.WorksheetRowIndex.Should().Be(2U);
        exception.SourceValue.Should().Be(4.4);
    }

    [Fact(Skip = "複数のプロパティが同じ列を指定した場合の検証を実装するときに解除する。")]
    public void 型付きTableは複数のプロパティが同じ列を指定した場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<DuplicateColumnRow>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(typeof(DuplicateColumnRow));
        exception.ColumnName.Should().Be("int");
        exception.WorksheetRowIndex.Should().BeNull();
        exception.SourceValue.Should().BeNull();
    }

    [Fact(Skip = "属性を付けたプロパティにpublic setterがない場合の検証を実装するときに解除する。")]
    public void 型付きTableは属性を付けたプロパティにpublicなsetterがない場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<AttributedPropertyWithoutPublicSetterRow>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(
            typeof(AttributedPropertyWithoutPublicSetterRow));
        exception.ColumnName.Should().Be("int");
        exception.PropertyName.Should().Be(
            nameof(
                AttributedPropertyWithoutPublicSetterRow.IntegerValue));
        exception.PropertyType.Should().Be(typeof(int));
        exception.WorksheetRowIndex.Should().BeNull();
        exception.SourceValue.Should().BeNull();
    }

    [Fact(Skip = "属性のない読み取り専用プロパティをマッピング対象外にする規則を実装するときに解除する。")]
    public void 型付きTableは属性のない読み取り専用プロパティを無視します()
    {
        var rows = book
            .ReadTable<RowWithReadOnlyProperty>(MappingTableName)
            .ToArray();

        rows
            .Select(row => row.IntegerValue)
            .Should()
            .Equal(4, 2, 3, 10, 1, 9, 6, 8, 5, 5, 7);

        rows
            .Select(row => row.Description)
            .Should()
            .OnlyContain(value => value == "computed");
    }

    [Fact(Skip = "publicな引数なしコンストラクターを要求する型生成規則を実装するときに解除する。")]
    public void 型付きTableはpublicな引数なしコンストラクターがない型では失敗します()
    {
        var action = () =>
        {
            book.ReadTable<RowWithoutPublicParameterlessConstructor>(MappingTableName)
                .ToArray();
        };

        var exception = action
            .Should()
            .Throw<TableMappingException>()
            .Which;

        exception.TableName.Should().Be(MappingTableName);
        exception.MappingType.Should().Be(
            typeof(RowWithoutPublicParameterlessConstructor));
        exception.ColumnName.Should().BeNull();
        exception.PropertyName.Should().BeNull();
        exception.PropertyType.Should().BeNull();
        exception.WorksheetRowIndex.Should().BeNull();
        exception.SourceValue.Should().BeNull();
    }

    public sealed class PropertyNameMappedRow
    {
        public int @int { get; set; }

        public double @float { get; set; }

        public string @string { get; set; } = "";
    }

    public sealed class IntegerOnlyRow
    {
        [SpreadsheetColumn("int")]
        public int IntegerValue { get; set; }
    }

    public sealed class MissingColumnRow
    {
        [SpreadsheetColumn("missing")]
        public int Value { get; set; }
    }

    public sealed class StringAsIntegerRow
    {
        [SpreadsheetColumn("string")]
        public int Value { get; set; }
    }

    public sealed class FloatingPointAsIntegerRow
    {
        [SpreadsheetColumn("float")]
        public int Value { get; set; }
    }

    public sealed class DuplicateColumnRow
    {
        [SpreadsheetColumn("int")]
        public int FirstValue { get; set; }

        [SpreadsheetColumn("int")]
        public int SecondValue { get; set; }
    }

    public sealed class AttributedPropertyWithoutPublicSetterRow
    {
        [SpreadsheetColumn("int")]
        public int IntegerValue { get; private set; }
    }

    public sealed class RowWithReadOnlyProperty
    {
        [SpreadsheetColumn("int")]
        public int IntegerValue { get; set; }

        public string Description => "computed";
    }

    public sealed class RowWithoutPublicParameterlessConstructor
    {
        public RowWithoutPublicParameterlessConstructor(int value)
        {
            Value = value;
        }

        [SpreadsheetColumn("int")]
        public int Value { get; set; }
    }
}
