using ClosedXML.Excel;
using ClosedXML.Report;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordTableExtractor.Core;
using WordTableExtractor.Features.Import;

namespace WordTableExtractor.Features.Export
{
    public class TableExporter
    {
        private ExportOptions _options;
        private TableImporter _importer;

        public TableExporter(ExportOptions options)
        {
            _options = options;
        }

        public void Export()
        {
            Console.WriteLine($"String import of '{_options.Filename}'.");

            _importer = new TableImporter(new ImportOptions
            {
                Filename = _options.Filename,
                Sheet = _options.Sheet,
                Range = _options.Range,
                Output = "",
                Consistency = false
            });

            var tree = _importer.ImportSpecification();

            Console.WriteLine($"Imported {tree.Root.Count} items as tree structure.");

            var nodes = GetFlatTree(tree);

            Console.WriteLine($"Flattened the to to a list of {nodes.Count} items.");

            var template = new XLWorkbook(_options.Template);
            var ws = template.Worksheets.First();

            switch(ws.Name)
            {
                case "Inbox":
                    WriteInboxTemplate(ws, nodes);
                    break;
            }

            template.SaveAs(_options.Output);

            Console.WriteLine($"Saved the filled template to '{_options.Output}'.");
        }

        private void WriteInboxTemplate(IXLWorksheet worksheet, List<TreeNode<SpecificationItem>> nodes)
        {
            var numberOfColumns = worksheet.Row(1).CellsUsed().Count();
            var firstRow = 2;

            Console.WriteLine($"Opened '{_options.Template}' as template with {numberOfColumns} used Columns.");
            Console.WriteLine($"Starting to write data at row {firstRow}.");

            for(int i = 0; i < nodes.Count; i++)
            {
                var node = nodes[i];

                // Initialize list of strings with empty values
                var rowData = new List<string>();
                for (int j = 0; j < numberOfColumns; j++)
                    rowData.Add(string.Empty);

                // Map data from Import to Export format
                // Type Folder, Information, Non-functional
                rowData[1] = node.Item.IsHeading ? "Folder" : node.Item.IsRequirement ? "Non-functional" : "Information";
                // Stakeholder EV-24 oder "EV-24, EV-46"
                rowData[2] = "EV-24";
                // Business Value Must have, Should have, Nice to have
                rowData[3] = "Must have";
                // Summary
                rowData[5] = node.Item.Summary;
                // Component
                rowData[9] = "";
                // Status n/a, New, Draft, Waiting for review, Rejected, Outdated (cancelled), In rating, Partly Accepted, Rework necessary, Rated
                rowData[10] = "";
                // Assigned to
                rowData[14] = "q271043,q271752";
                // Team
                rowData[15] = "";
                // Lead Team
                rowData[16] = "";
                // Description
                rowData[29] = node.Item.Content;

                // Set Data
                for(int j = 0; j < numberOfColumns; j++)
                {
                    if(!string.IsNullOrEmpty(rowData[j]))
                        worksheet.Cell(firstRow + i, j + 1).Value = rowData[j];
                }

                // Set indent for structure on Summary
                worksheet.Cell(firstRow + i, 6).Style.Alignment.Indent = node.Level;
            }            
        }

        public void FillInboxReportTemplate()
        {
            Console.WriteLine($"String import of '{_options.Filename}'.");

            _importer = new TableImporter(new ImportOptions
            {
                Filename = _options.Filename,
                Sheet = _options.Sheet,
                Range = _options.Range,
                Output = "",
                Consistency = false
            });

            var tree = _importer.ImportSpecification();

            Console.WriteLine($"Imported {tree.Root.Count} items as tree structure.");

            var flatTree = GetFlatTree(tree);
            var nodes = flatTree.Select(x => x.Item).ToList();

            Console.WriteLine($"Flattened the to to a list of {nodes.Count} items.");

            var template = new XLTemplate(_options.Template);
            Console.WriteLine($"Opened '{_options.Template}' as template.");

            Console.WriteLine("Adding variables to the template.");
            template.AddVariable(new { Nodes = nodes });

            Console.WriteLine("Generating the template.");
            template.Generate();

            // Set indentation for summary to keep structure
            Console.WriteLine("Starting to set indentation for Summary column to keep structure.");

            var ws = template.Workbook.Worksheets.First(x => !x.Name.StartsWith("_"));
            var usedRows = ws.RowsUsed().Count();
            var minIndent = nodes.Min(x => x.Level);

            for(int i = 2; i <= usedRows; i++)
            {
                if(ws.Cell(i, 2).TryGetValue<int>(out int indent))
                {
                    ws.Cell(i, 7).Style.Alignment.Indent = indent - minIndent;
                }

                ws.Cell(i, 2).Value = string.Empty;
            }

            template.SaveAs(_options.Output);

            Console.WriteLine($"Saved the filled template to '{_options.Output}'.");
        }

        private List<TreeNode<T>> GetFlatTree<T>(Tree<T> tree) where T: ITreeNode
        {
            var nodes = GetTreeNodes(tree.Root);

            return nodes;
        }

        private List<TreeNode<T>> GetTreeNodes<T>(TreeNode<T> treeNode) where T: ITreeNode
        {
            var nodes = new List<TreeNode<T>>();
            
            if (treeNode.Item != null)
                nodes.Add(treeNode);


            foreach(var child in treeNode.Children)
            {
                var childNodes = GetTreeNodes(child);

                if (childNodes.Count > 0)
                    nodes.AddRange(childNodes);
            }

            return nodes;
        }
    }
}
