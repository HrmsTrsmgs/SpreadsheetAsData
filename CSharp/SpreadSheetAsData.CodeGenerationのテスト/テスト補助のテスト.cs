using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class テスト補助のテスト
{
    [Fact]
    public void TypeNamesは生成ソースに含まれる型名を返します()
    {
        CodeGenerationSpec
            .FromSources(
                """
                namespace Generated;

                public class BasicStructureBook
                {
                }

                public class SalesDataSheet
                {
                }
                """)
            .TypeNames
            .Should()
            .BeEquivalentTo("BasicStructureBook", "SalesDataSheet");
    }

    [Fact]
    public void GeneratedTypeは指定した型の名前を返します()
    {
        CodeGenerationSpec
            .FromSources(
                """
                namespace Generated;

                public class SalesDetailTable
                {
                }
                """)
            .GeneratedType("SalesDetailTable")
            .Name
            .Should()
            .Be("SalesDetailTable");
    }

    [Fact]
    public void BaseTypeNameはジェネリック型引数を含む基底型名を返します()
    {
        CodeGenerationSpec
            .FromSources(
                """
                namespace Generated;

                public class SalesDetailTable : Table<SalesDetail>
                {
                }

                public class SalesDetail
                {
                }
                """)
            .GeneratedType("SalesDetailTable")
            .BaseTypeName
            .Should()
            .Be("Table<SalesDetail>");
    }

    [Fact]
    public void PropertyNamesは指定した型に含まれるプロパティ名を返します()
    {
        CodeGenerationSpec
            .FromSources(
                """
                namespace Generated;

                public class SalesDetail
                {
                    public int CustomerId { get; set; }
                    public double Amount { get; set; }
                }
                """)
            .GeneratedType("SalesDetail")
            .PropertyNames
            .Should()
            .BeEquivalentTo("CustomerId", "Amount");
    }
}
