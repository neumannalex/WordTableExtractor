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

            var dataTable = new DataTable();
            dataTable.TableName = ws.Name;

            var columnNames = hasHeader ? range.Row(1).Cells().Select(x => x.GetValue<string>()).ToList() :
                range.Row(1).Cells().Select(x => x.Address.ColumnLetter).ToList();

            foreach (var columnName in columnNames)
                dataTable.Columns.Add(columnName);

            var firstRow = hasHeader ? 2 : 1;
            for(int i = firstRow; i <= range.RowCount(); i++)
            {
                var row = dataTable.NewRow();
                var data = range.Row(i).Cells().Select(x => x.HasFormula ? x.CachedValue.ToString() : x.GetValue<string>()).ToArray();
                row.ItemArray = data;

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        public NeumannAlex.Tree.Tree<Dictionary<string, string>> ImportAsTree(string filename, string sheet, string dataRange, string hierarchyColumn, string typeColumn, string directoryType, bool hasHeader)
        {
            var wb = new XLWorkbook(filename);

            var ws = wb.Worksheet(sheet);

            var range = ws.Range(dataRange);            

            var columnNames = hasHeader ? range.Row(1).Cells().Select(x => x.GetValue<string>()).ToList() :
                range.Row(1).Cells().Select(x => x.Address.ColumnLetter).ToList();

            var firstRow = hasHeader ? 2 : 1;

            var tree = new NeumannAlex.Tree.Tree<Dictionary<string, string>>();
            tree.Tags["Name"] = ws.Name;

            NeumannAlex.Tree.ITreeNode<Dictionary<string, string>> lastNode = tree.TreeRoot;
            int lastDepth = 0;

            for (int i = firstRow; i <= range.RowCount(); i++)
            {
                var values = range.Row(i).Cells().Select(x => x.HasFormula ? x.CachedValue.ToString().Trim() : x.GetValue<string>().Trim()).ToArray();
                var dict = columnNames.Zip(values, (k, v) => new { k, v }).ToDictionary(x => x.k, x => (string)x.v);

                var node = new SectionizedTreeNode<Dictionary<string, string>>(dict)
                {
                    Type = dict[typeColumn] == directoryType ? SectionizedTreeNodeType.Folder : SectionizedTreeNodeType.Leaf
                };


                if(!string.IsNullOrEmpty(hierarchyColumn) && columnNames.Contains(hierarchyColumn))
                {
                    var hierarchy = dict[hierarchyColumn];
                    var depth = CalculateDepthFromHierarchy(hierarchy);

                    if(depth > lastDepth)
                    {
                        lastNode = lastNode.AddChild(node);
                    }
                    else if(depth == lastDepth)
                    {
                        lastNode = lastNode.Parent.AddChild(node);
                    }
                    else
                    {
                        var depthOffset = lastNode.Depth - lastDepth;
                        var ancestorDepth = depth + depthOffset - 1;
                        
                        var p = tree.LastOrDefault(x => x.Depth == ancestorDepth);
                        if (p != null)
                            lastNode = p.AddChild(node);
                        else
                            throw new IndexOutOfRangeException($"Could not find last node at level {depth - 1}.");
                    }

                    lastDepth = depth;
                }
                else
                {
                    tree.AddChild(node);
                }
            }

            return tree;
        }

        private int CalculateDepthFromHierarchy(string hierarchy)
        {
            hierarchy = hierarchy.Replace('-', '.');

            if(hierarchy.Contains('.'))
            {
                var parts = hierarchy.Split('.');
                return parts.Length;
            }
            else
            {
                return 1;
            }
        }

        public static string FormatTreenodeForDump(NeumannAlex.Tree.ITreeNode<Dictionary<string, string>> node)
        {
            if (node.IsRoot)
                return "Root";

            var path = node.PathString;

            int maxLength = 30;

            var contents = node.Value["Contents"].Length < maxLength ? node.Value["Contents"] : node.Value["Contents"].Substring(0, maxLength - 3) + "...";

            return $"[{path}] [{node.Value["fixedsection"]}] - {node.Value["Artifact Type"]} - {contents}";
        }
    }
}
