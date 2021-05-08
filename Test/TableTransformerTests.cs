using FluentAssertions;
using System;
using System.Collections.Generic;
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
            Filename = @"D:\Temp\lh\Example\Example.xlsx",
            Sheet = "tbl",
            Range = "A1:G622",
            Output = @"D:\Temp\lh\Example\Example_transformed.xlsx",
            Consistency = true
        };

        [Fact]
        public void RunTransformation()
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

        [Fact]
        public void TestHierarchy()
        {
            var h = new Hierarchy();

            h.HasLevel(1).Should().BeTrue();
            h.HasLevel(2).Should().BeFalse();

            h.ToString().Should().Be("1");

            h.NewLevel();
            h.HasLevel(2).Should().BeTrue();
            h.ToString().Should().Be("1.1");

            h.IncrementLevel(2);
            h.ToString().Should().Be("1.2");
        }

        [Fact]
        public void TestTree()
        {
            var req0 = new RequirementNode("1.15", "Root", "");

            var rootNode = new TreeNode<RequirementNode>(req0);

            

        }
    }
}
