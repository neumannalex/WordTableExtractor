using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using WordTableExtractor;
using WordTableExtractor.Import;
using Xunit;

namespace Test
{
    public class TableTransformerTests
    {
        private ImportOptions _options = new ImportOptions
        {
            Filename = @"D:\Temp\lh\Example\Example-imported-numbered.xlsx",
            Sheet = "tbl",
            Range = "A1:J621",
            Output = @"D:\Temp\lh\Example\Example-transformed-test.xlsx",
            Consistency = true
        };

        [Fact]
        public void RunTransformation()
        {
            var transformer = new TableImporter(_options);

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

        [Fact]
        public void TestTree()
        {
            //var req0 = new RequirementNode("1.15", "Root", "");

            //var rootNode = new TreeNode<RequirementNode>(req0);

            

        }
    }
}
