using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class テスト補助のテスト
{
    [Fact]
    public void TypeNamesは生成ソースに含まれる型名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
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
    public void GeneratedTypeは指定した名前の生成型を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
                """
                namespace Generated;

                public class BasicStructureBook
                {
                }

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
    public void Nameは生成型の名前を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
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
    public void NamespaceNameは指定した生成型の名前空間を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
                """
                namespace Generated.Custom;

                public class SalesDataSheet
                {
                }
                """)
            .GeneratedType("SalesDataSheet")
            .NamespaceName
            .Should()
            .Be("Generated.Custom");
    }

    [Fact]
    public void ToStringは生成型の名前空間付き表示名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
                """
                namespace Generated;

                public class SalesDataSheet
                {
                }
                """)
            .GeneratedType("SalesDataSheet")
            .ToString()
            .Should()
            .Be("SalesDataSheet: Generated");
    }

    [Fact]
    public void BaseTypeNameはジェネリック型引数を含む基底型名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFromSources(
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
        GeneratedCodeInspection
            .SyntaxFromSources(
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
