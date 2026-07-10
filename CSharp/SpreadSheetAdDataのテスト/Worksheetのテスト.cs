using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAdData.Test
{
    public class Worksheetのテスト : IDisposable
    {

        Workbook book;
        Worksheet sheet1;
        Worksheet data;

        public Worksheetのテスト()
        {
            book = Workbook.Open(@"TestData\Book1.xlsx");
            sheet1 = book.Sheets["Sheet1"];
            data = book.Sheets["いろいろなデータ"];
        }

        public void Dispose()
        {
            book.Close();
        }

        [Fact]
        public void Nameはシート名を返します()
        {
            Assert.Equal("Sheet1", sheet1.Name);
        }

        [Fact]
        public void Nameは日本語で指定したシート名も適切に扱います()
        {
            Assert.Equal("いろいろなデータ", data.Name);
        }

        [Fact]
        public void Bookは所属しているWorkbookを取得します()
        {
            Assert.Same(book, sheet1.Book);
        }

        [Fact]
        public void Cellsはセルの参照文字列を文字列を指定してセル取得します()
        {
            Assert.Equal("C3", sheet1.Cells["C3"].Reference);
        }

        [Fact]
        public void Cellsはセルの参照文字列に存在しないセル名を指定した時にFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(()=> sheet1.Cells["a1"]);
        }

        [Fact]
        public void Cellsは一文字の列名称の列を取得します()
        {
            Assert.Equal("A1", sheet1.Cells["A1"].Reference);
            Assert.Equal("Z1", sheet1.Cells["Z1"].Reference);
        }

        [Fact]
        public void Cellsは二文字の列名称の列を取得します()
        {
            Assert.Equal("AA1", sheet1.Cells["AA1"].Reference);
            Assert.Equal("ZZ1", sheet1.Cells["ZZ1"].Reference);
        }

        [Fact]
        public void Cellsは三文字の列名称の列を取得します()
        {
            Assert.Equal("AAA1", sheet1.Cells["AAA1"].Reference);
            Assert.Equal("XFD1", sheet1.Cells["XFD1"].Reference);
        }

        [Fact]
        public void Cellsは大きすぎる列名を指定した時にFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => sheet1.Cells["XFE1"]);
            Assert.ThrowsAny<FormatException>(() => sheet1.Cells["ZZZ1"]);
            Assert.ThrowsAny<FormatException>(() => sheet1.Cells["AAAA1"]);
        }

        [Fact]
        public void Cellsは大きな行番号のセルを取得します()
        {
            Assert.Equal("A1048576", sheet1.Cells["A1048576"].Reference);
        }

        [Fact]
        public void Cellsは大きすぎる行番号を指定した時にFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => sheet1.Cells["A1048577"]);
        }

        [Fact]
        public void Cellsは空のセルも取得します()
        {
            Assert.Equal("B1", sheet1.Cells["B1"].Reference);
        }

        [Fact]
        public void Cellsはは同じセルの場合は同じオブジェクトを取得します()
        {
            Assert.Same(sheet1.Cells["A1"], sheet1.Cells["A1"]);
        }

        [Fact]
        public void Cellsはセルの座標を数値で指定してセル取得します()
        {
            Assert.Equal("C3", sheet1.Cells[3, 3].Reference);
        }

        [Fact]
        public void Cellsは指定方法が違っても同じセルの場合は同じオブジェクトを取得します()
        {
            Assert.Same(sheet1.Cells["A1"], sheet1.Cells[1, 1]);
        }

        [Fact]
        public void Rangeは2引数を指定して範囲を取得します()
        {
            Assert.Equal("A1:C3", sheet1.Range["A1", "C3"].ToString());
            Assert.Equal("B2:B2", sheet1.Range["B2", "B2"].ToString());
        }

        [Fact]
        public void Rangeは2引数を指定して同じ範囲を指定した場合に同じセルを返します()
        {
            Assert.Same(sheet1.Range["A1", "C3"], sheet1.Range["A1", "C3"]);
        }
    }
}
