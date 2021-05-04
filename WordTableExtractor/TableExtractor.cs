using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTableExtractor
{
    public class TableExtractor
    {
        private readonly Options _options;
        private List<DataTable> _dataTables;
        private WordprocessingDocument _document;
        private Summary _summary;

        public string OutputFolderPath
        {
            get
            {
                var filename = System.IO.Path.GetFileNameWithoutExtension(_options.Output);
                return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(_options.Output), filename);
            }
        }

        public string OutputFilePath
        {
            get
            {
                var filename = System.IO.Path.GetFileNameWithoutExtension(_options.Output);
                return System.IO.Path.Combine(OutputFolderPath, filename + ".xlsx");
            }
        }

        public string ImageFolderPath
        {
            get
            {
                return System.IO.Path.Combine(OutputFolderPath, "images");
            }
        }

        public string ImageZipPath
        {
            get
            {
                return System.IO.Path.Combine(OutputFolderPath, "images.zip");
            }
        }


        public TableExtractor(Options options)
        {
            _options = options;

            // Default for output filename and directory
            if (string.IsNullOrEmpty(_options.Output))
                options.Output = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(_options.Filename), System.IO.Path.GetFileNameWithoutExtension(_options.Filename) + ".xlsx");


            _dataTables = new List<DataTable>();
        }


        public Summary ExtractTables()
        {
            InformationMessage("Starting to export tables from Word document.");

            if (OpenWordDocument())
            {
                InformationMessage("Importing tables from Word document.");
                ImportTablesFromWord();

                InformationMessage("Exporting tables to Excel document.");
                ExportTablesToExcel();

                InformationMessage("Postprocessing exported images.");
                HandleExportedImages();
            }

            InformationMessage("Export finished.");

            return _summary;
        }

        private bool OpenWordDocument()
        {
            _summary.InputFilePath = _options.Filename;
            _summary.ShouldExportImages = _options.ExportImages;
            _summary.ShouldZipExportedImages = _options.ZipImages;
            _summary.OutputFilePath = OutputFilePath;

            try
            {
                _document = WordprocessingDocument.Open(_options.Filename, false);

                _summary.CouldReadWordDocument = true;

                VerboseMessage("Loading Word document succeeded.");

                return true;
            }
            catch(Exception ex)
            {
                _document = null;

                _summary.CouldReadWordDocument = false;

                VerboseMessage("Loading Word document failed.");

                return false;
            }
        }        

        private void ImportTablesFromWord()
        {
            var tables = _document.MainDocumentPart.Document.Body.Elements<Table>();

            _summary.NumberOfTablesFound = tables.Count();

            VerboseMessage($"Found {tables.Count()} tables in Word document.");

            var counter = 0;

            foreach (var table in tables)
            {
                VerboseMessage($"Importing table {++counter}/{tables.Count()}.");

                var dataTable = ExtractTable(table);

                var newIndex = _dataTables.Count + 1;

                dataTable.TableName = table.LocalName ?? $"Table {newIndex}";

                _dataTables.Add(dataTable);

                _summary.NumberOfTablesImported++;
            }
        }

        private void ExportTablesToExcel()
        {
            using(var workbook = new XLWorkbook())
            {
                var counter = 0;

                foreach(var table in _dataTables)
                {
                    try
                    {
                        VerboseMessage($"Exporting table {++counter}/{_dataTables.Count}.");

                        workbook.Worksheets.Add(table);
                        _summary.NumberOfTablesImported++;

                        VerboseMessage($"Exporting table {counter} succeeded.");
                    }
                    catch
                    {
                        VerboseMessage($"Exporting table {counter} failed.");
                        _summary.Errors.Add($"Could not create worksheet for table '{table.TableName}'.");
                    }
                }

                try
                {
                    workbook.SaveAs(OutputFilePath);
                    _summary.CouldWriteExcelDocument = true;

                    VerboseMessage($"Saving Excel File to '{OutputFilePath}' succeeded.");
                }
                catch(Exception ex)
                {
                    VerboseMessage($"Saving Excel File to '{OutputFilePath}' failed.");

                    _summary.Errors.Add($"Could not save Excel file because of '{ex.Message}'.");
                    _summary.CouldWriteExcelDocument = false;
                }
            }
        }

        private DataTable ExtractTable(Table table)
        {
            DataTable dataTable = new DataTable();
            int rowCount = 0;

            IEnumerable<TableRow> rows = table.Elements<TableRow>();

            foreach (TableRow row in rows)
            {
                if (rowCount == 0)
                {
                    foreach (TableCell cell in row.Descendants<TableCell>())
                    {
                        dataTable.Columns.Add(cell.InnerText);
                    }
                    rowCount += 1;
                }
                else
                {
                    dataTable.Rows.Add();
                    int i = 0;
                    foreach (TableCell cell in row.Descendants<TableCell>())
                    {
                        var content = GetCellContents(cell);

                        dataTable.Rows[dataTable.Rows.Count - 1][i] = content;
                        i++;
                    }
                }
            }

            return dataTable;
        }

        private string GetCellContents(TableCell cell)
        {
            var content = cell.InnerText;

            if(_options.ExportImages)
            {
                var imageNames = ExportImagesFromCell(cell);

                var sb = new StringBuilder();
                
                sb.AppendLine(cell.InnerText);

                foreach (var imageName in imageNames)
                    sb.AppendLine($"<{imageName}>");

                content = sb.ToString();
            }


            return content;
        }

        private List<string> ExportImagesFromCell(TableCell cell)
        {
            var imageNames = new List<string>();

            foreach(var par in cell.Descendants<Paragraph>())
            {
                foreach(var run in par.Descendants<Run>())
                {
                    foreach(var pic in run.Descendants<ImageData>())
                    {
                        _summary.NumberOfImagesFound++;

                        foreach(var imagedata in pic.Descendants<ImageData>())
                        {
                            if(!string.IsNullOrEmpty(imagedata.RelationshipId))
                            {
                                try
                                {
                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}'.");

                                    ImagePart imagePart = (ImagePart)_document.MainDocumentPart.GetPartById(imagedata.RelationshipId);
                                    Image img = Image.FromStream(imagePart.GetStream());

                                    var filename = imagePart.Uri.OriginalString.Substring(imagePart.Uri.OriginalString.LastIndexOf("/") + 1);

                                    if (!System.IO.Directory.Exists(ImageFolderPath))
                                        System.IO.Directory.CreateDirectory(ImageFolderPath);

                                    img.Save(System.IO.Path.Combine(ImageFolderPath, filename));

                                    imageNames.Add(filename);

                                    _summary.NumberOfImagesExported++;

                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}' succeeded.");
                                }
                                catch(Exception ex)
                                {
                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}' failed.");

                                    _summary.Errors.Add($"Could not export image for RelationshipId '{imagedata.RelationshipId}' because of '{ex.Message}'.");
                                }
                            }
                        }
                    }
                }
            }

            return imageNames;
        }        

        private void HandleExportedImages()
        {
            if(System.IO.Directory.Exists(ImageFolderPath))
            {
                var numImages = System.IO.Directory.GetFiles(ImageFolderPath).Length;

                if(numImages > 0)
                {
                    if(_options.ZipImages)
                    {
                        try
                        {
                            // put all images in a zip file and delete image folder
                            ZipFile.CreateFromDirectory(ImageFolderPath, ImageZipPath);

                            VerboseMessage("Zipping exported images succeeded.");
                        }
                        catch(Exception ex)
                        {
                            VerboseMessage("Zipping exported images failed.");
                            _summary.Errors.Add($"Could not create Zip-File because of '{ex.Message}'.");
                        }
                    }
                }
                else
                {
                    VerboseMessage("No images were exported but image output directory exists. Deleting directory.");
                    System.IO.Directory.Delete(ImageFolderPath);
                }
            }
        }

        private void VerboseMessage(string message)
        {
            if (_options.Verbose)
                Console.WriteLine(message);
        }

        private void InformationMessage(string message)
        {
            Console.WriteLine(message);
        }
    }
}
