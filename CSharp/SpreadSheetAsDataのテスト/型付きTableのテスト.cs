using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class 型付きTableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";
    const string MappingTableName = "型付き行マッピング";
    const string PropertyNameMappingTableName = "プロパティ名マッピング";

    readonly Workbook book;
    readonly Table<TestMappedRow> tested;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public 型付きTableのテスト()
    {
        book = Workbook.Open(TestFilePath);
        tested = book.ReadTable<TestMappedRow>(MappingTableName);
    }

    public void Dispose()
    {
        book.Close();
        temporaryFiles.Dispose();
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
    public void 型付きTableは真偽値をboolプロパティへ読み込みます()
    {
        book.ReadTable<BooleanOnlyRow>(MappingTableName)
            .Select(it => it.BooleanValue)
            .Should().Equal(false, true, false);
    }

    [Fact]
    public void 型付きTableは数値をnullableなdoubleプロパティへ読み込みます()
    {
        book.ReadTable<NullableDoubleRow>(MappingTableName)
            .Select(it => it.FloatingPointValue)
            .Should().Equal(4.4, 5.5, 6.6);
    }

    [Fact]
    public void 型付きTableは空白セルをnullableなdoubleプロパティのnullとして読み込みます()
    {
        book.ReadTable<NullableDoubleRow>("空白数値マッピング")
            .Select(it => it.FloatingPointValue)
            .Should().Equal(0d, null);
    }

    [Fact]
    public void 型付きTableは整数値をnullableなintプロパティへ読み込みます()
    {
        book.ReadTable<NullableIntegerRow>(MappingTableName)
            .Select(it => it.IntegerValue)
            .Should().Equal(1, 2, 3);
    }

    [Fact]
    public void 型付きTableは空白セルをnullableなintプロパティのnullとして読み込みます()
    {
        book.ReadTable<NullableIntegerWithBlankRow>("空白数値マッピング")
            .Select(it => it.IntegerValue)
            .Should().Equal(0, null);
    }

    [Fact]
    public void 型付きTableは真偽値をnullableなboolプロパティへ読み込みます()
    {
        book.ReadTable<NullableBooleanRow>(MappingTableName)
            .Select(it => it.BooleanValue)
            .Should().Equal(false, true, false);
    }

    [Fact]
    public void 型付きTableは空白セルをnullableなboolプロパティのnullとして読み込みます()
    {
        book.ReadTable<NullableBooleanRow>("空白数値マッピング")
            .Select(it => it.BooleanValue)
            .Should().Equal(false, null);
    }

    [Fact]
    public void 型付きTableは空白セルをstringプロパティの空文字列として読み込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));
        book.Tables[MappingTableName].Rows.First()["文字列"].Value = null;

        book.ReadTable<StringOnlyRow>(MappingTableName)
            .Select(it => it.TextValue)
            .Should().Equal("", "たちつてと", "なにぬねの");
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
    public void Replaceは属性で指定した列の値を置き換えます()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.ReadTable<WritableMappedRow>(MappingTableName)
                .Replace(
                    [
                        new WritableMappedRow
                        {
                            IntegerValue = 10,
                            FloatingPointValue = 1.5,
                            TextValue = "first"
                        },
                        new WritableMappedRow
                        {
                            IntegerValue = 20,
                            FloatingPointValue = 2.5,
                            TextValue = "second"
                        },
                        new WritableMappedRow
                        {
                            IntegerValue = 30,
                            FloatingPointValue = 3.5,
                            TextValue = "third"
                        }
                    ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.ReadTable<TestMappedRow>(MappingTableName)
            .Should().BeEquivalentTo(
                [
                    new TestMappedRow
                    {
                        IntegerValue = 10,
                        FloatingPointValue = 1.5,
                        TextValue = "first"
                    },
                    new TestMappedRow
                    {
                        IntegerValue = 20,
                        FloatingPointValue = 2.5,
                        TextValue = "second"
                    },
                    new TestMappedRow
                    {
                        IntegerValue = 30,
                        FloatingPointValue = 3.5,
                        TextValue = "third"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void Replaceはnullableなdoubleの値とnullを数値と空白として書き込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        book.ReadTable<NullableDoubleRow>(MappingTableName)
            .Replace(
                [
                    new NullableDoubleRow { FloatingPointValue = 1.5 },
                    new NullableDoubleRow { FloatingPointValue = null },
                    new NullableDoubleRow { FloatingPointValue = 3.5 }
                ]);

        book.Tables[MappingTableName].Rows
            .Select(it => it["数値"].Value as object)
            .Should().Equal(1.5, new BlankValue(), 3.5);
    }

    [Fact]
    public void Replaceはnullableなintの値とnullを数値と空白として書き込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        book.ReadTable<NullableIntegerRow>(MappingTableName)
            .Replace(
                [
                    new NullableIntegerRow { IntegerValue = 10 },
                    new NullableIntegerRow { IntegerValue = null },
                    new NullableIntegerRow { IntegerValue = 30 }
                ]);

        book.Tables[MappingTableName].Rows
            .Select(it => it["数値2"].Value as object)
            .Should().Equal(10d, new BlankValue(), 30d);
    }

    [Fact]
    public void Replaceはnullableなboolの値とnullを真偽値と空白として書き込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        book.ReadTable<NullableBooleanRow>(MappingTableName)
            .Replace(
                [
                    new NullableBooleanRow { BooleanValue = true },
                    new NullableBooleanRow { BooleanValue = null },
                    new NullableBooleanRow { BooleanValue = false }
                ]);

        book.Tables[MappingTableName].Rows
            .Select(it => it["真偽値"].Value as object)
            .Should().Equal(true, new BlankValue(), false);
    }

    [Fact]
    public void Replaceは文字列プロパティの空文字列を空白セルとして書き込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        book.ReadTable<StringOnlyRow>(MappingTableName)
            .Replace(
                [
                    new StringOnlyRow { TextValue = "" },
                    new StringOnlyRow { TextValue = " " },
                    new StringOnlyRow { TextValue = "text" }
                ]);

        book.Tables[MappingTableName].Rows
            .Select(it => it["文字列"].Value as object)
            .Should().Equal(new BlankValue(), " ", "text");
    }

    [Fact]
    public void Replaceは属性がないプロパティ名を列名として使用します()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.ReadTable<WritablePropertyNameMappedRow>(PropertyNameMappingTableName)
                .Replace(
                    [
                        new WritablePropertyNameMappedRow
                        {
                            @float = 10.1,
                            @string = "first"
                        },
                        new WritablePropertyNameMappedRow
                        {
                            @float = 20.2,
                            @string = "second"
                        }
                    ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.ReadTable<PropertyNameMappedRow>(PropertyNameMappingTableName)
            .Should().BeEquivalentTo(
                [
                    new PropertyNameMappedRow
                    {
                        @float = 10.1,
                        @string = "first"
                    },
                    new PropertyNameMappedRow
                    {
                        @float = 20.2,
                        @string = "second"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void Replaceは書き込み元に対応プロパティがない列を変更しません()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.ReadTable<WritableIntegerOnlyRow>(MappingTableName)
                .Replace(
                    [
                        new WritableIntegerOnlyRow { IntegerValue = 10 },
                        new WritableIntegerOnlyRow { IntegerValue = 20 },
                        new WritableIntegerOnlyRow { IntegerValue = 30 }
                    ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);
        var rows = tested.ReadTable<TestMappedRow>(MappingTableName).ToArray();

        rows.Select(it => it.IntegerValue).Should().Equal(10, 20, 30);
        rows.Select(it => it.TextValue).Should().Equal("さしすせそ", "たちつてと", "なにぬねの");
    }

    [Fact]
    public void Replaceはワークシート上の順序でデータ行を置き換えます()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.ReadTable<WritableIntegerOnlyRow>(MappingTableName)
                .Replace(
                    [
                        new WritableIntegerOnlyRow { IntegerValue = 10 },
                        new WritableIntegerOnlyRow { IntegerValue = 20 },
                        new WritableIntegerOnlyRow { IntegerValue = 30 }
                    ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.Tables[MappingTableName]
            .Rows
            .Select(it => it["数値2"].Value as object)
            .Should().Equal(10d, 20d, 30d);
    }

    [Fact]
    public void Replaceは属性で指定した列が存在しない場合に失敗します()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            book.ReadTable<WritableMissingColumnRow>(MappingTableName)
                .Replace([new WritableMissingColumnRow { Value = 1 }]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact]
    public void Replaceは複数プロパティが同じ列を指定した場合に失敗します()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            book.ReadTable<WritableDuplicateColumnRow>(MappingTableName)
                .Replace(
                    [
                        new WritableDuplicateColumnRow
                        {
                            FirstValue = 1,
                            SecondValue = 2
                        }
                    ]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact]
    public void Replaceは属性を付けたプロパティにpublicなgetterがない場合に失敗します()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            book.ReadTable<WritableAttributedPropertyWithoutPublicGetterRow>(MappingTableName)
                .Replace([new WritableAttributedPropertyWithoutPublicGetterRow()]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact]
    public void Replaceは属性のない書き込み専用プロパティを無視します()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.ReadTable<WritableRowWithWriteOnlyProperty>(MappingTableName)
                .Replace([new WritableRowWithWriteOnlyProperty { IntegerValue = 10 }]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Tables[MappingTableName].Rows.First()["数値2"].Value as object)
            .Should().Be(10d);
    }

    [Fact]
    public void Replaceはセル値へ変換できないプロパティ型の場合に失敗します()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            book.ReadTable<WritableUnsupportedValueRow>(MappingTableName)
                .Replace([new WritableUnsupportedValueRow { Value = DateTime.Today }]);
        };

        action.Should().Throw<TableMappingException>();
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
    public void 型付きTableは小数をnullableなintへ変換しようとした場合に失敗します()
    {
        var action = () =>
        {
            book.ReadTable<NullableIntegerWithBlankRow>(MappingTableName)
                .ToArray();
        };

        var thrown = action.Should().Throw<TableMappingException>();

        thrown.Which.PropertyType.Should().Be(typeof(int?));
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

    public sealed class WritableMappedRow
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; set; }

        [SpreadSheetName("数値")]
        public double FloatingPointValue { get; set; }

        [SpreadSheetName("文字列")]
        public string TextValue { get; set; } = "";
    }

    public sealed class WritablePropertyNameMappedRow
    {
        public double @float { get; set; }

        public string @string { get; set; } = "";
    }

    public sealed class WritableIntegerOnlyRow
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; set; }
    }

    public sealed class WritableMissingColumnRow
    {
        [SpreadSheetName("missing")]
        public int Value { get; set; }
    }

    public sealed class WritableDuplicateColumnRow
    {
        [SpreadSheetName("数値2")]
        public int FirstValue { get; set; }

        [SpreadSheetName("数値2")]
        public int SecondValue { get; set; }
    }

    public sealed class WritableAttributedPropertyWithoutPublicGetterRow
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue
        {
            set { }
        }
    }

    public sealed class WritableRowWithWriteOnlyProperty
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; set; }

        public string Description
        {
            set { }
        }
    }

    public sealed class WritableUnsupportedValueRow
    {
        [SpreadSheetName("数値2")]
        public DateTime Value { get; set; }
    }

    public sealed class StringOnlyRow
    {
        [SpreadSheetName("文字列")]
        public string TextValue { get; set; } = "";
    }

    public sealed class NullableDoubleRow
    {
        [SpreadSheetName("数値")]
        public double? FloatingPointValue { get; set; }
    }

    public sealed class NullableIntegerRow
    {
        [SpreadSheetName("数値2")]
        public int? IntegerValue { get; set; }
    }

    public sealed class NullableIntegerWithBlankRow
    {
        [SpreadSheetName("数値")]
        public int? IntegerValue { get; set; }
    }

    public sealed class NullableBooleanRow
    {
        [SpreadSheetName("真偽値")]
        public bool? BooleanValue { get; set; }
    }

    public sealed class BooleanOnlyRow
    {
        [SpreadSheetName("真偽値")]
        public bool BooleanValue { get; set; }
    }

    public sealed class IntegerOnlyRow
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; set; }
    }

    public sealed class MissingColumnRow
    {
        [SpreadSheetName("missing")]
        public int Value { get; set; }
    }

    public sealed class StringAsIntegerRow
    {
        [SpreadSheetName("文字列")]
        public int Value { get; set; }
    }

    public sealed class FloatingPointAsIntegerRow
    {
        [SpreadSheetName("数値")]
        public int Value { get; set; }
    }

    public sealed class EmptyTableRow
    {
        [SpreadSheetName("列1")]
        public string FirstValue { get; set; } = "";

        [SpreadSheetName("列2")]
        public string SecondValue { get; set; } = "";
    }

    public sealed class DuplicateColumnRow
    {
        [SpreadSheetName("数値2")]
        public int FirstValue { get; set; }

        [SpreadSheetName("数値2")]
        public int SecondValue { get; set; }
    }

    public sealed class AttributedPropertyWithoutPublicSetterRow
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; private set; }
    }

    public sealed class RowWithReadOnlyProperty
    {
        [SpreadSheetName("数値2")]
        public int IntegerValue { get; set; }

        public string Description => "computed";
    }

    public sealed class RowWithoutPublicParameterlessConstructor
    {
        public RowWithoutPublicParameterlessConstructor(int value)
        {
            Value = value;
        }

        [SpreadSheetName("数値2")]
        public int Value { get; set; }
    }
}
