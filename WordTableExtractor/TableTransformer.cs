using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordTableExtractor
{
    public class TableTransformer
    {
        private TransformOptions _options;
        private List<RequirementNode> _nodes;

        public TableTransformer(TransformOptions options)
        {
            _options = options;
        }

        public void Transform()
        {
            var importedData = ImportData();

            BuildNodes(importedData);

            if(_options.Consistency)
                CheckConsistency();

            var processedData = ProcessTable(importedData);

            ExportData(processedData);
        }

        public DataTable ImportData()
        {
            var wb = new XLWorkbook(_options.Filename);

            var ws = wb.Worksheet(_options.Sheet);

            var range = ws.Range(_options.Range);

            var dataTable = new DataTable();
            dataTable.TableName = ws.Name;

            // Column names
            var firstRow = range.Row(1);
            foreach(var cell in firstRow.Cells())
            {
                dataTable.Columns.Add((string)cell.Value);
            }

            dataTable.Columns.Add("originalsection");

            // Data
            for (int i = 2; i < range.Rows().Count(); i++)
            {
                var values = range.Row(i).Cells().Select(x => x.Value).ToArray();
                var row = dataTable.NewRow();
                row.ItemArray = values;
                
                // Workaround für falsche Section
                // Es gibt viele Einträge, deren Section a.b.0-x anstelle von a.b-x lautet
                // Fälle, wo anstelle der 0 eine andere Ziffer steht kommen nicht vor
                // Maßnahme: ersetze .0-x durch -x
                var originalSection = row.Field<string>("section");
                var correctedSection = Regex.Replace(originalSection, @"\.0-(\d+)$", "-$1");

                row["originalsection"] = originalSection;
                row["section"] = correctedSection;

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        public void BuildNodes(DataTable table)
        {
            _nodes = new List<RequirementNode>();

            // Read all nodes
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var rowData = table.Rows[i];

                var node = new RequirementNode(rowData.Field<string>("section"), rowData.Field<string>("Artifact Type"), rowData.Field<string>("Contents"));
                _nodes.Add(node);
            }
        }

        public void CheckConsistency()
        {
            var path = System.IO.Path.GetDirectoryName(_options.Filename);
            var filename = System.IO.Path.GetFileNameWithoutExtension(_options.Filename);
            var logfilename = System.IO.Path.Combine(path, $"{filename}-consistency.xlsx");

            var logTable = new DataTable();
            logTable.TableName = "Log";

            logTable.Columns.Add("Heading item");
            logTable.Columns.Add("Next item");
            logTable.Columns.Add("Heading level");
            logTable.Columns.Add("Next level");

            for (int i = 0; i < _nodes.Count - 1; i++)
            {
                if(_nodes[i].IsHeading)
                {
                    var nextNode = _nodes[i + 1];

                    var currentLevel = _nodes[i].Level;
                    var nextLevel = nextNode.Level;

                    var row = logTable.NewRow();

                    var logItems = new List<object>{
                        _nodes[i].Address.ToString(),
                        nextNode.Address.ToString(),
                        currentLevel,
                        nextLevel
                    };

                    row.ItemArray = logItems.ToArray();
                    logTable.Rows.Add(row);
                }
            }

            var wb = new XLWorkbook();
            wb.Worksheets.Add(logTable);
            wb.SaveAs(logfilename);
        }

        private void ExportData(DataTable dataTable)
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(dataTable);
                workbook.SaveAs(_options.Output);
            }
        }

        private DataTable ProcessTable(DataTable table)
        {
            // Loop through all rows
            table.Columns.Add("chapter");
            table.Columns.Add("title");
            table.Columns.Add("level");
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var rowData = table.Rows[i];

                var node = new RequirementNode(rowData.Field<string>("section"), rowData.Field<string>("Artifact Type"), rowData.Field<string>("Contents"));
                var parentNode = _nodes.Where(x => x.Address == node.Address.ParentAddress).FirstOrDefault();

                if(parentNode == null)
                {
                    Console.WriteLine("Parent Node is NULL");
                }

                if (node.IsHeading)
                {
                    rowData["chapter"] = parentNode == null ? node.Title : parentNode.Title;
                }
                else
                {
                    rowData["chapter"] = parentNode == null ? "Uuups" : parentNode.Title;
                }

                rowData["title"] = node.Title;
                rowData["level"] = node.Level;
            }

            return table;
        }
    }
}
