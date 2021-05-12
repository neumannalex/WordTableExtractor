using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordTableExtractor.Features.Import;

namespace WordTableExtractor.Features.Sanitize
{
    public class Sanitizer
    {
        private SanitizerOptions _options;

        public Sanitizer(SanitizerOptions options)
        {
            _options = options;
        }

        public void SanitizeNumbering()
        {
            var importer = new TableImporter(new ImportOptions
            {
                Filename = _options.Filename,
                Output = _options.Output,
                Sheet = _options.Sheet,
                Range = _options.Range,
                Consistency = false
            });

            var tree = importer.ImportSpecification();

            var nodes = tree.ToList();

            var wb = new XLWorkbook(_options.Filename);
            var ws = wb.Worksheet(_options.Sheet);
            var range = ws.Range(_options.Range);
            var hasHeader = true;

            var columnNames = hasHeader ? range.Row(1).CellsUsed().Select(x => x.GetValue<string>()).ToList() :
                    range.Cells().Select(x => x.Address.ColumnLetter).ToList();

            var numberingColumnName = _options.Column;

            if (!columnNames.Contains(numberingColumnName))
            {
                Console.WriteLine($"No column named '{numberingColumnName}' found in the header.");
                return;
            }

            var numberingColumnIndex = columnNames.IndexOf(numberingColumnName) + 1;            

            for(int i = 2; i <= range.RowCount(); i++)
            {
                var oldNumbering = range.Cell(i, numberingColumnIndex).GetValue<string>();

                var node = nodes.Where(x => x.Item.Address.ToString() == oldNumbering).FirstOrDefault();
                if(node != null)
                {
                    range.Cell(i, numberingColumnIndex).SetValue<string>(node.Numbering);
                }
            }

            wb.SaveAs(_options.Output);
        }

    }
}
