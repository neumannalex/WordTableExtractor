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
    public class CommonTests
    {
        [Fact]
        public void TestReportgenerator()
        {
            var generator = new ReportGenerator();
            generator.Run();
        }
    }
}
