using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using WordTableExtractor;
using WordTableExtractor.Features.Import;
using Xunit;

namespace Test
{
    public class TableTransformerTests
    {
        private ImportOptions _options = new ImportOptions
        {
            Filename = "",
            Sheet = "tbl",
            Range = "A1:J621",
            Output = "",
            Consistency = false
        };

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
