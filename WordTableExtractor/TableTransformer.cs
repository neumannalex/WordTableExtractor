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

            var columnNames = firstRow.Cells().Select(x => (string)x.Value).ToList();
            
            foreach(var name in columnNames)
                dataTable.Columns.Add(name);

            dataTable.Columns.Add("originalsection");
            dataTable.Columns.Add("sourcerange");

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

                row["sourcerange"] = range.Row(i).RangeAddress.ToString();
                row["originalsection"] = originalSection;
                row["section"] = correctedSection;

                dataTable.Rows.Add(row);

                var node = new RequirementNode(originalSection,
                    row.Field<string>("Artifact Type"),
                    row.Field<string>("Contents"),
                    range.Row(i).RangeAddress.ToString()
                    );

                node.Values = columnNames.Zip(values, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
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

                var node = new RequirementNode(rowData.Field<string>("section"), rowData.Field<string>("Artifact Type"), rowData.Field<string>("Contents"), rowData.Field<string>("sourcerange"));
                _nodes.Add(node);
            }
        }

        public void CheckConsistency()
        {
            var path = System.IO.Path.GetDirectoryName(_options.Filename);
            var filename = System.IO.Path.GetFileNameWithoutExtension(_options.Filename);
            var logfilename = System.IO.Path.Combine(path, $"{filename}-consistency.xlsx");

            var levelChangeTable = BuildTableOfItemLevelChangesForConsistencyCheck();
            var hierarchyTable = BuildHierarchieTableForConsistencyCheck();

            var wb = new XLWorkbook();
            wb.Worksheets.Add(levelChangeTable);
            wb.Worksheets.Add(hierarchyTable);
            wb.SaveAs(logfilename);
        }

        private DataTable BuildTableOfItemLevelChangesForConsistencyCheck()
        {
            var table = new DataTable();
            table.TableName = "Level Changes";

            table.Columns.Add("Current Item section");
            table.Columns.Add("Current Item title");
            table.Columns.Add("Current Item type");
            table.Columns.Add("Current Item level");

            table.Columns.Add("Next Item section");
            table.Columns.Add("Next Item title");
            table.Columns.Add("Next Item type");
            table.Columns.Add("Next Item level");

            table.Columns.Add("Level difference");

            for (int i = 0; i < _nodes.Count - 1; i++)
            {
                if (_nodes[i].IsHeading)
                {
                    var currentNode = _nodes[i];
                    var nextNode = _nodes[i + 1];

                    var row = table.NewRow();

                    row["Current Item section"] = currentNode.Section;
                    row["Current Item title"] = currentNode.Title;
                    row["Current Item type"] = currentNode.Type.ToString();
                    row["Current Item level"] = currentNode.Level;

                    row["Next Item section"] = nextNode.Section;
                    row["Next Item title"] = nextNode.Title;
                    row["Next Item type"] = nextNode.Type.ToString();
                    row["Next Item level"] = nextNode.Level;

                    row["Level difference"] = nextNode.Level - currentNode.Level;

                    table.Rows.Add(row);
                }
            }

            return table;
        }

        private DataTable BuildHierarchieTableForConsistencyCheck()
        {
            var table = new DataTable();
            table.TableName = "Hierarchy";

            var minLevel = _nodes.Min(x => x.Level);
            var maxLevel = _nodes.Max(x => x.Level);
            
            // Add columns
            for (int level = minLevel; level <= maxLevel; level++)
                table.Columns.Add($"Level {level}");

            table.Columns.Add("Type");
            table.Columns.Add("Section");
            table.Columns.Add("Source");
            table.Columns.Add("Content");
            table.Columns.Add("Chapter");

            var minHierarchyLevel = 2;
            var lastLevel = minHierarchyLevel;
            var currentLevel = minHierarchyLevel;

            //Tree<RequirementNode> tree = new Tree<RequirementNode>();
            //TreeNode<RequirementNode> currentTreeNode = null;
            //TreeNode<RequirementNode> lastTreeNode = tree.Root;

            foreach (var node in _nodes)
            {
                var levelColName = $"Level {node.Level}";

                var row = table.NewRow();

                row[levelColName] = node.Title ?? "<Empty>";
                row["Type"] = node.Type.ToString() ?? "<Empty>";
                row["Section"] = node.Section ?? "<Empty>";
                row["Source"] = node.SourceRange;

                var content = "<Empty>";
                if(!string.IsNullOrEmpty(node.Content))
                    content = node.Content.Length > 50 ? node.Content.Substring(0, 50) + "..." : node.Content;

                row["Content"] = content;

                // Chapter handling
                //currentLevel = node.Level;

                //if(currentLevel > lastLevel)
                //{
                //    currentTreeNode = lastTreeNode.AddChild(node);
                //}
                //else if(currentLevel == lastLevel)
                //{
                //    if(lastTreeNode.IsRoot)
                //    {
                //        currentTreeNode = lastTreeNode.AddChild(node);
                //    }
                //    else
                //    {
                //        if(node.IsHeading)
                //            currentTreeNode = lastTreeNode.Parent.AddChild(node);
                //        else
                //            lastTreeNode.Parent.AddChild(node);
                //    }
                //}
                //else
                //{
                //    if(!lastTreeNode.Parent.IsRoot)
                //        currentTreeNode = lastTreeNode.Parent.Parent.AddChild(node);
                //}

                //lastLevel = currentLevel;
                //lastTreeNode = currentTreeNode;

                //row["Chapter"] = hierarchy.ToString();

                table.Rows.Add(row);
            }


            //var x = PrintTreeNode(tree.Root);

            return table;
        }

        private string PrintTreeNode(TreeNode<RequirementNode> node, int level = 0)
        {
            string indent = level > 0 ? " ".Repeat(level) : string.Empty;

            var sb = new StringBuilder();

            if(!node.IsRoot)
            {
                sb.AppendLine($"{indent}{node.Item.Section} ({node.Level}) [{node.Chapter}]");
            }
                

            foreach(var child in node.Children)
            {
                sb.AppendLine(PrintTreeNode(child, level + 1));
            }

            return sb.ToString();
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
