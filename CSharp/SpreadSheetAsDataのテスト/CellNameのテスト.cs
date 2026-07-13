using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class CellNameのテスト
{
    const string セル参照のゼロ行列検証保留理由 =
        "CellNameが1始まりの行番号と列番号を保証する実装時に解除する。";

    [Fact]
    public void ParseメソッドでCellNameが生成できます()
    {
        FluentActions.Invoking(
            () => CellName.Parse("A1")
        ).Should().NotThrow();
    }

    [Fact]
    public void 列番号と行番号を指定してCellNameがせいせいできます()
    {
        FluentActions.Invoking(
            () => new CellName(1, 1)
        ).Should().NotThrow();
    }

    [Fact]
    public void 同じセル位置は同一として扱われます()
    {
        CellName.Parse("A1").Should().Be(CellName.Parse("A1"));
    }

    [Fact]
    public void 行が違うと同一ではないとして扱われます()
    {
        CellName.Parse("A2").Should().NotBe(CellName.Parse("A1"));
    }

    [Fact]
    public void 列が違うと同一ではないとして扱われます()
    {
        CellName.Parse("B1").Should().NotBe(CellName.Parse("A1"));
    }

    [Fact]
    public void 指定したセルの位置を一意としてハッシュのキーとして使えます()
    {
        var set = new HashSet<CellName>
        {
            CellName.Parse("A1")
        };
        set.Should().ContainSingle();
        set.Add(CellName.Parse("A1"));
        set.Should().ContainSingle();
        set.Add(CellName.Parse("B1"));
        set.Should().HaveCount(2);
        set.Add(CellName.Parse("A2"));
        set.Should().HaveCount(3);
    }

    [Fact]
    public void Parseメソッドに無効なセル名を指定するとFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => CellName.Parse("A")
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => CellName.Parse("1")
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => CellName.Parse("A1A1")
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void Parseは行番号が0のセル参照を変換しません()
    {
        var action = () => CellName.Parse("A0");

        action.Should().Throw<FormatException>();
    }

    [Fact]
    public void Parseメソッドに大きすぎる列名を指定した場合はFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => CellName.Parse("XFE1")
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => CellName.Parse("AAAA1")
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => CellName.Parse("ZZZZ1")
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void コンストラクタに大きすぎる列番号を指定した場合はFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => new CellName(16385, 1)
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void コンストラクターは列番号が0の場合に失敗します()
    {
        var action = () => new CellName(0, 1);

        action.Should().Throw<FormatException>();
    }

    [Fact]
    public void Parseメソッドに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => CellName.Parse("A1048577")
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void コンストラクタに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => new CellName(1, 1048577)
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void コンストラクターは行番号が0の場合に失敗します()
    {
        var action = () => new CellName(1, 0);

        action.Should().Throw<FormatException>();
    }

    [Fact]
    public void ColumnNameプロパティで列名を取得できます()
    {
        CellName.Parse("A1").ColumnName.Should().Be("A");
        CellName.Parse("B1").ColumnName.Should().Be("B");
    }

    [Fact]
    public void 行列番号で生成した場合にColumnNameプロパティで列名を取得できます()
    {
        new CellName(1, 1).ColumnName.Should().Be("A");
        new CellName(2, 1).ColumnName.Should().Be("B");
    }

    [Fact]
    public void ColumnIndexプロパティで一文字の列番号を取得できます()
    {
        CellName.Parse("A1").ColumnIndex.Should().Be(1U);
        CellName.Parse("Z1").ColumnIndex.Should().Be(26U);
    }

    [Fact]
    public void ColumnIndexプロパティで二文字のの列番号を取得できます()
    {
        CellName.Parse("AA1").ColumnIndex.Should().Be(27U);
        CellName.Parse("ZZ1").ColumnIndex.Should().Be(702U);
    }

    [Fact]
    public void ColumnIndexプロパティで三文字のの列番号を取得できます()
    {
        CellName.Parse("AAA1").ColumnIndex.Should().Be(703U);
        CellName.Parse("XFD1").ColumnIndex.Should().Be(16384U);
    }

    [Fact]
    public void コンストラクタで生成した場合にColumnIndexプロパティで大きな行番号を取得できます()
    {
        new CellName(16384, 1).ColumnIndex.Should().Be(16384U);
    }

    [Fact]
    public void RowIndexプロパティで行番号を取得できます()
    {
        CellName.Parse("A1").RowIndex.Should().Be(1U);
        CellName.Parse("A2").RowIndex.Should().Be(2U);
    }

    [Fact]
    public void 行列番号で生成した場合にRowIndexプロパティで行番号を取得できます()
    {
        new CellName(1, 1).RowIndex.Should().Be(1U);
        new CellName(1, 2).RowIndex.Should().Be(2U);
    }

    [Fact]
    public void Parseメソッドで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
    {
        CellName.Parse("A1048576").RowIndex.Should().Be(1048576U);
    }

    [Fact]
    public void コンストラクタで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
    {
        new CellName(1, 1048576).RowIndex.Should().Be(1048576U);
    }

    [Fact]
    public void ToStringメソッドはA1形式で文字列を返します()
    {
        CellName.Parse("A1").ToString().Should().Be("A1");
        CellName.Parse("B2").ToString().Should().Be("B2");
        CellName.Parse("AA1").ToString().Should().Be("AA1");
        CellName.Parse("AAA1").ToString().Should().Be("AAA1");
        CellName.Parse("XFD1048576").ToString().Should().Be("XFD1048576");
    }
}
