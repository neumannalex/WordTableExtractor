using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using WordTableExtractor;
using WordTableExtractor.Features.Export;
using WordTableExtractor.Features.Extract;
using WordTableExtractor.Features.Import;
using WordTableExtractor.Features.Sanitize;
using Xunit;

namespace Test
{
    public class CommonTests
    {
        [Fact]
        public void RunExtract()
        {
            var options = new ExtractOptions
            {
                Filename = "",
                Output = "",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);
            extractor.ExtractTables();
        }

        [Fact]
        public void RunSanitizer()
        {
            var options = new SanitizerOptions
            {
                Filename = "",
                Sheet = "tbl",
                Range = "A1:J621",
                Output = "",
                Column = "section"
            };

            var sanitizer = new Sanitizer(options);
            sanitizer.SanitizeNumbering();
        }

        [Fact]
        public void RunStructure()
        {
            var  options = new ImportOptions
            {
                Filename = "",
                Sheet = "tbl",
                Range = "A1:J621",
                Output = "",
                Consistency = false
            };

            var transformer = new TableImporter(options);

            transformer.AnalyzeStructure();
        }

        [Fact]
        public void RunExport()
        {
            var options = new ExportOptions
            {
                Filename = "",
                Sheet = "tbl",
                Range = "A1:J621",
                Output = "",
                Template = ""
            };

            var exporter = new TableExporter(options);
            exporter.FillInboxReportTemplate();
        }
    }
}
