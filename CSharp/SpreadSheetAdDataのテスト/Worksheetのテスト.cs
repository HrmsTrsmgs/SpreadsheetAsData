using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using NUnit.Framework.Constraints;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class Worksheetのテスト
    {

        Workbook book;
        Worksheet sheet1;
        Worksheet data;

        [SetUp]
        public void SetUp()
        {
            book = Workbook.Open(@"TestData\Book1.xlsx");
            sheet1 = book.Sheets["Sheet1"];
            data = book.Sheets["いろいろなデータ"];
        }

        [TearDown]
        public void TearDown()
        {
            book.Close();
        }

        [Test]
        public void Nameはシート名を返します()
        {
            Assert.That(sheet1.Name, Is.EqualTo("Sheet1"));
        }

        [Test]
        public void Nameは日本語で指定したシート名も適切に扱います()
        {
            Assert.That(data.Name, Is.EqualTo("いろいろなデータ"));
        }

        [Test]
        public void Bookは所属しているWorkbookを取得します()
        {
            Assert.That(sheet1.Book, Is.SameAs(book));
        }

        [Test]
        public void Cellsはセルの参照文字列を文字列を指定してセル取得します()
        {
            Assert.That(sheet1.Cells["C3"].Reference, Is.EqualTo("C3"));
        }

        [Test]
        public void Cellsはセルの参照文字列に存在しないセル名を指定した時にFormatExceptionを投げます()
        {
            Assert.That(()=> sheet1.Cells["a1"], Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Cellsは一文字の列名称の列を取得します()
        {
            Assert.That(sheet1.Cells["A1"].Reference, Is.EqualTo("A1"));
            Assert.That(sheet1.Cells["Z1"].Reference, Is.EqualTo("Z1"));
        }

        [Test]
        public void Cellsは二文字の列名称の列を取得します()
        {
            Assert.That(sheet1.Cells["AA1"].Reference, Is.EqualTo("AA1"));
            Assert.That(sheet1.Cells["ZZ1"].Reference, Is.EqualTo("ZZ1"));
        }

        [Test]
        public void Cellsは三文字の列名称の列を取得します()
        {
            Assert.That(sheet1.Cells["AAA1"].Reference, Is.EqualTo("AAA1"));
            Assert.That(sheet1.Cells["XFD1"].Reference, Is.EqualTo("XFD1"));
        }

        [Test]
        public void Cellsは大きすぎる列名を指定した時にFormatExceptionを投げます()
        {
            Assert.That(() => sheet1.Cells["XFE1"], Throws.InstanceOf<FormatException>());
            Assert.That(() => sheet1.Cells["ZZZ1"], Throws.InstanceOf<FormatException>());
            Assert.That(() => sheet1.Cells["AAAA1"], Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Cellsは大きな行番号のセルを取得します()
        {
            Assert.That(sheet1.Cells["A1048576"].Reference, Is.EqualTo("A1048576"));
        }

        [Test]
        public void Cellsは大きすぎる行番号を指定した時にFormatExceptionを投げます()
        {
            Assert.That(() => sheet1.Cells["A1048577"], Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Cellsは空のセルも取得します()
        {
            Assert.That(sheet1.Cells["B1"].Reference, Is.EqualTo("B1"));
        }

        [Test]
        public void Cellsはは同じセルの場合は同じオブジェクトを取得します()
        {
            Assert.That(sheet1.Cells["A1"], Is.SameAs(sheet1.Cells["A1"]));
        }

        [Test]
        public void Cellsはセルの座標を数値で指定してセル取得します()
        {
            Assert.That(sheet1.Cells[3, 3].Reference, Is.EqualTo("C3"));
        }

        [Test]
        public void Cellsは指定方法が違っても同じセルの場合は同じオブジェクトを取得します()
        {
            Assert.That(sheet1.Cells[1, 1], Is.SameAs(sheet1.Cells["A1"]));
        }

        [Test]
        public void Rangeは2引数を指定して範囲を取得します()
        {
            Assert.That(sheet1.Range["A1", "C3"].ToString(), Is.EqualTo("A1:C3"));
            Assert.That(sheet1.Range["B2", "B2"].ToString(), Is.EqualTo("B2:B2"));
        }

        [Test]
        public void Rangeは2引数を指定して同じ範囲を指定した場合に同じセルを返します()
        {
            Assert.That(sheet1.Range["A1", "C3"], Is.SameAs(sheet1.Range["A1", "C3"]));
        }
    }
}
