using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using System.IO;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class Cellのテスト
    {
        Worksheet いろいろなデータ;
        Cell a1;
        Cell b1;
        Cell a2;
        Cell b2;
        Cell a3;
        Cell b3;

        [SetUp]
        public void SetUp()
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
        [TearDown]
        public void TearDown()
        {
            いろいろなデータ.Book.Close();
        }

        [Test]
        public void Bookプロパティはブックを取得できます()
        {
            Assert.That(a1.Book, Is.SameAs(いろいろなデータ.Book));
        }

        [Test]
        public void Sheetプロパティはシートを取得できます()
        {
            Assert.That(a1.Sheet, Is.SameAs(いろいろなデータ));
        }

        [Test]
        public void Valueプロパティは数字の値を取得できます()
        {
            Assert.That(a1.Value, Is.EqualTo(1.1));
            Assert.That(b1.Value, Is.EqualTo(2.2));
        }

        [Test]
        public void Valueプロパティはboolの値を取得できます()
        {
            
            Assert.That(a2.Value, Is.True);
            Assert.That(b2.Value, Is.False);
        }

        [Test]
        public void Valueプロパティは文字列の値を取得できます()
        {
            Assert.That(a3.Value, Is.EqualTo("あいうえお"));
            Assert.That(b3.Value, Is.EqualTo("かきくけこ"));
        }

        [Test]
        public void Referenceプロパティがセル参照の名称を取得できます()
        {
            Assert.That(a1.Reference, Is.EqualTo("A1"));
            Assert.That(b1.Reference, Is.EqualTo("B1"));
        }
        [Test]
        public void RowIndexプロパティが行番号を取得できます()
        {
            Assert.That(a1.RowIndex, Is.EqualTo(1));
            Assert.That(a2.RowIndex, Is.EqualTo(2));
        }

        [Test]
        public void ColumnIndexプロパティが列番号を取得できます()
        {
            Assert.That(a1.ColumnIndex, Is.EqualTo(1));
            Assert.That(b1.ColumnIndex, Is.EqualTo(2));
        }
    }
}
