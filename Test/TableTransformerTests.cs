using FluentAssertions;
using System;
using System.Data;
using System.Text.RegularExpressions;
using WordTableExtractor;
using Xunit;

namespace Test
{
    public class TableTransformerTests
    {
        private TransformOptions _options = new TransformOptions
        {
            Filename = @"C:\Daten\Temp\LH_Word2Excel\20210503_HWComponentSpecification\20210503_HWComponentSpecification.xlsx",
            Sheet = "tbl",
            Range = "A1:G644",
            Output = @"C:\Daten\Temp\LH_Word2Excel\20210503_HWComponentSpecification\20210503_HWComponentSpecification_transformed.xlsx",
            Consistency = true
        };

        [Fact]
        public void ImportData()
        {
            var transformer = new TableTransformer(_options);

            transformer.Transform();
        }

        [Fact]
        public void TestRegex()
        {
            var input = "1.15.1.0-1";
            var pattern = @"\.0-(\d+)$";

            var result = Regex.Match(input, pattern);

            var replaced = Regex.Replace(input, pattern, "-$1");
        }
    }
}
