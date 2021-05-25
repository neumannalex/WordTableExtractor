using FluentAssertions;
using System;
using WordTableExtractor;
using WordTableExtractor.Services;
using Xunit;
namespace Test
{
    public class ExcelTableImporterTests
    {
        private readonly string _filename = @"D:\Temp\lh\Example\Example-imported.xlsx";
        private readonly string _sheet = "tbl";
        private readonly string _range = "A1:J621";
        private readonly string _hierarchyColumn = "fixedsection";
        private readonly string _typeColumn = "Artifact Type";
        private readonly string _directoryType = "Heading";
        private readonly bool _captions = true;

        [Fact]
        public void ImportAsTable()
        {
            var importer = new ExcelTableImporter();

            var table = importer.ImportAsTable(_filename, _sheet, _range, _captions);

            table.Should().NotBeNull();
            table.TableName.Should().Be(_sheet);
            table.Rows.Should().HaveCount(620);
            table.Columns.Should().HaveCount(10);
        }

        [Fact]
        public void ImportAsTree()
        {
            var importer = new ExcelTableImporter();

            var tree = importer.ImportAsTree(_filename, _sheet, _range, _hierarchyColumn, _typeColumn, _directoryType, _captions);

            tree.Should().NotBeNull();
            tree.Count.Should().Be(620);

            var dump = tree.TreeRoot.ToText("", ExcelTableImporter.FormatTreenodeForDump);
        }
    }
}
