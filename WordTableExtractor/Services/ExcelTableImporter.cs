using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Data;
using WordTableExtractor.Core;

namespace WordTableExtractor.Services
{
    public class ExcelTableImporter
    {
        public DataTable ImportAsTable(string filename, string sheet, string dataRange, bool hasHeader)
        {
            var wb = new XLWorkbook(filename);

            var ws = wb.Worksheet(sheet);

            var range = ws.Range(dataRange);

            var columnNames = hasHeader ? range.Row(1).Cells().Select(x => x.GetValue<string>()).ToList() :
                range.Row(1).Cells().Select(x => x.Address.ColumnLetter).ToList();

            var dataTable = new DataTable();

            dataTable.TableName = ws.Name;

            var firstRow = hasHeader ? 2 : 1;
            for(int i = firstRow; i < range.RowCount(); i++)
            {
                var row = dataTable.NewRow();
                row.ItemArray = range.Row(i).Cells().Select(x => x.Value).ToArray();

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        public Tree<SpecificationItem> ImportAsTree(string filename, string sheet, string dataRange, bool hasHeader)
        {
            return null;
        }
    }
}
