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
    public class Workbookのテスト : IDisposable
    {

        Workbook book1;

        public const string コピーパス = @"TestData\Book1-Copy.xlsx";

        public Workbookのテスト()
        {
            if (!File.Exists(コピーパス))
            {
                File.Copy(@"TestData\Book1.xlsx", コピーパス);
            }
            book1 = Workbook.Open(@"TestData\Book1.xlsx");
        }

        public void Dispose()
        {
            book1.Close();
        }
        
        [Fact]
        public void Openはファイルを束縛します()
        {
            var tested = Workbook.Open(コピーパス);
            try
            {
                Assert.ThrowsAny<IOException>(() => File.Delete(コピーパス));
            }
            finally
            {
                tested.Close();
            }
        }

        [Fact]
        public void Closeはファイルの束縛を解除します()
        {
            var tested = Workbook.Open(コピーパス);

            tested.Close();
            Assert.Null(Record.Exception(() => File.Delete(コピーパス)));
        }

        [Fact]
        public void Disposeはファイルの束縛を解除します()
        {
            using (var tested = Workbook.Open(コピーパス))
            {
            }
            Assert.Null(Record.Exception(() => File.Delete(コピーパス)));
        }

        [Fact]
        public void Sheetsでシートが取得できます()
        {
            Assert.Equal(3, book1.Sheets.Count);
        }

        [Fact]
        public void Sheetsに数字を指定してシートが取得できます()
        {
            Assert.Equal("Sheet1", book1.Sheets[0].Name);
        }

        [Fact]
        public void Sheetsにシート名を指定してシートが取得できます()
        {
            Assert.Equal("Sheet1", book1.Sheets["Sheet1"].Name);
        }

        [Fact]
        public void インデクサに数字を指定してシートが取得できます()
        {
            Assert.Equal("Sheet1", book1[0].Name);
        }

        [Fact]
        public void インデクサにシート名を指定してシートが取得できます()
        {
            Assert.Equal("Sheet1", book1["Sheet1"].Name);
        }

    }
}
