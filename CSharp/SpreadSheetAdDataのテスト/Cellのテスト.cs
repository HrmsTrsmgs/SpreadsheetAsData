using Marimo.SpreadSheetAsData;
using System;
using Xunit;
using System.IO;

namespace Marimo.SpreadSheetAdData.Test
{
    public class Cellのテスト : IDisposable
    {
        Worksheet いろいろなデータ;
        Cell a1;
        Cell b1;
        Cell a2;
        Cell b2;
        Cell a3;
        Cell b3;

        public Cellのテスト()
        {
            var book = Workbook.Open(@"TestData\Book1.xlsx");
            
            いろいろなデータ = book.Sheets["いろいろなデータ"];

            a1 = いろいろなデータ.Cells["A1"];
            b1 = いろいろなデータ.Cells["B1"];
            a2 = いろいろなデータ.Cells["A2"];
            b2 = いろいろなデータ.Cells["B2"];
            a3 = いろいろなデータ.Cells["A3"];
            b3 = いろいろなデータ.Cells["B3"];
            
        }
        public void Dispose()
        {
            いろいろなデータ.Book.Close();
        }

        [Fact]
        public void Bookプロパティはブックを取得できます()
        {
            Assert.Same(いろいろなデータ.Book, a1.Book);
        }

        [Fact]
        public void Sheetプロパティはシートを取得できます()
        {
            Assert.Same(いろいろなデータ, a1.Sheet);
        }

        [Fact]
        public void Valueプロパティは数字の値を取得できます()
        {
            Assert.Equal(1.1, a1.Value);
            Assert.Equal(2.2, b1.Value);
        }

        [Fact]
        public void Valueプロパティはboolの値を取得できます()
        {
            
            Assert.True((bool)(a2.Value));
            Assert.False((bool)(b2.Value));
        }

        [Fact]
        public void Valueプロパティは文字列の値を取得できます()
        {
            Assert.Equal("あいうえお", a3.Value);
            Assert.Equal("かきくけこ", b3.Value);
        }

        [Fact]
        public void Referenceプロパティがセル参照の名称を取得できます()
        {
            Assert.Equal("A1", a1.Reference);
            Assert.Equal("B1", b1.Reference);
        }
        [Fact]
        public void RowIndexプロパティが行番号を取得できます()
        {
            Assert.Equal((uint)1, a1.RowIndex);
            Assert.Equal((uint)2, a2.RowIndex);
        }

        [Fact]
        public void ColumnIndexプロパティが列番号を取得できます()
        {
            Assert.Equal((uint)1, a1.ColumnIndex);
            Assert.Equal((uint)2, b1.ColumnIndex);
        }
    }
}
