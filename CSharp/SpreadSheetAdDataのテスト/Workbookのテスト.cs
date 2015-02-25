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
    public class Workbookのテスト : AssertionHelper
    {

        Workbook book1;

        public const string コピーパス = @"TestData\Book1-Copy.xlsx";

        [SetUp]
        public void SetUp()
        {
            if (!File.Exists(コピーパス))
            {
                File.Copy(@"TestData\Book1.xlsx", コピーパス);
            }
            book1 = Workbook.Open(@"TestData\Book1.xlsx");
        }

        [TearDown]
        public void TearDown()
        {
            book1.Close();
        }
        
        [Test]
        public void Openはファイルを束縛します()
        {
            var tested = Workbook.Open(コピーパス);
            Expect(() => File.Delete(コピーパス), Throws.InstanceOf<IOException>());
        }

        [Test]
        public void Closeはファイルの束縛を解除します()
        {
            var tested = Workbook.Open(コピーパス);

            tested.Close();
            Expect(() => File.Delete(コピーパス), Throws.Nothing);
        }

        [Test]
        public void Disposeはファイルの束縛を解除します()
        {
            using (var tested = Workbook.Open(コピーパス))
            {
            }
            Expect(() => File.Delete(コピーパス), Throws.Nothing);
        }

        [Test]
        public void Sheetsでシートが取得できます()
        {
            Expect(book1.Sheets.Count, Is.EqualTo(3));
        }

        [Test]
        public void Sheetsに数字を指定してシートが取得できます()
        {
            Expect(book1.Sheets[0].Name, Is.EqualTo("Sheet1"));
        }

        [Test]
        public void Sheetsにシート名を指定してシートが取得できます()
        {
            Expect(book1.Sheets["Sheet1"].Name, Is.EqualTo("Sheet1"));
        }

    }
}
