using Marimo.SpreadSheetAsData;
using NUnit.Framework;



namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class Cellのテスト : AssertionHelper
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
            Expect(a1.Book, Is.SameAs(いろいろなデータ.Book));
        }

        [Test]
        public void Sheetプロパティはシートを取得できます()
        {
            Expect(a1.Sheet, Is.SameAs(いろいろなデータ));
        }

        [Test]
        public void Valueプロパティは数字の値を取得できます()
        {
            Expect(a1.Value, Is.EqualTo(1.1));
            Expect(b1.Value, Is.EqualTo(2.2));
        }

        [Test]
        public void Valueプロパティはboolの値を取得できます()
        {
            Expect(a2.Value, Is.True);
            Expect(b2.Value, Is.False);
        }

        [Test]
        public void Valueプロパティは文字列の値を取得できます()
        {
            Expect(a3.Value, Is.EqualTo("あいうえお"));
            Expect(b3.Value, Is.EqualTo("かきくけこ"));
        }
    }
}
