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
using WordTableExtractor.Extensions;

namespace WordTableExtractor.Features.Extract
{
    public class TableExtractor
    {
        private readonly ExtractOptions _options;
        private List<DataTable> _dataTables;
        private WordprocessingDocument _document;
        private ExtractSummary _summary;

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


        public TableExtractor(ExtractOptions options)
        {
            _options = options;
            _summary = new ExtractSummary();

            // Default for output filename and directory
            if (string.IsNullOrEmpty(_options.Output))
                options.Output = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(_options.Filename), System.IO.Path.GetFileNameWithoutExtension(_options.Filename) + ".xlsx");


            _dataTables = new List<DataTable>();
        }


        public ExtractSummary ExtractTables()
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

            var extension = System.IO.Path.GetExtension(_options.Filename);
            if(extension.ToLower() != ".docx")
            {
                _summary.Errors.Add("Only *.docx files are supported.");
                VerboseMessage("Only *.docx files are supported.", 1);

                return false;
            }

            try
            {
                _document = WordprocessingDocument.Open(_options.Filename, false);

                _summary.CouldReadWordDocument = true;

                VerboseMessage("Loading Word document succeeded.", 1);

                return true;
            }
            catch(Exception ex)
            {
                _document = null;

                _summary.CouldReadWordDocument = false;

                _summary.Errors.Add($"Loading Word document failed because of '{ex.Message}'.");

                VerboseMessage($"Loading Word document failed because of '{ex.Message}'.", 1);

                return false;
            }
        }        

        private void ImportTablesFromWord()
        {
            var tables = _document.MainDocumentPart.Document.Body.Elements<Table>();

            _summary.NumberOfTablesFound = tables.Count();

            VerboseMessage($"Found {tables.Count()} tables in Word document.", 1);

            var counter = 0;

            foreach (var table in tables)
            {
                VerboseMessage($"Importing table {++counter}/{tables.Count()}.", 1);

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
                        VerboseMessage($"Exporting table {++counter}/{_dataTables.Count}.", 1);

                        workbook.Worksheets.Add(table);
                        _summary.NumberOfTablesExported++;

                        VerboseMessage($"Exporting table {counter} succeeded.", 1);
                    }
                    catch
                    {
                        VerboseMessage($"Exporting table {counter} failed.", 1);
                        _summary.Errors.Add($"Could not create worksheet for table '{table.TableName}'.");
                    }
                }

                try
                {
                    workbook.SaveAs(OutputFilePath);
                    _summary.CouldWriteExcelDocument = true;

                    VerboseMessage($"Saving Excel File to '{OutputFilePath}' succeeded.", 1);
                }
                catch(Exception ex)
                {
                    VerboseMessage($"Saving Excel File to '{OutputFilePath}' failed.", 1);

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
                        dataTable.Columns.Add(cell.InnerText.Trim());
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
            var content = cell.InnerText.Trim();

            if(_options.ExportImages)
            {
                var imageNames = ExportImagesFromCell(cell);

                var sb = new StringBuilder();
                
                sb.Append(content);

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
                    foreach(var pic in run.Descendants<Picture>())
                    {
                        _summary.NumberOfImagesFound++;

                        foreach(var imagedata in pic.Descendants<ImageData>())
                        {
                            if(!string.IsNullOrEmpty(imagedata.RelationshipId))
                            {
                                try
                                {
                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}'.", 2);

                                    ImagePart imagePart = (ImagePart)_document.MainDocumentPart.GetPartById(imagedata.RelationshipId);
                                    Image img = Image.FromStream(imagePart.GetStream());

                                    var filename = imagePart.Uri.OriginalString.Substring(imagePart.Uri.OriginalString.LastIndexOf("/") + 1);

                                    if (!System.IO.Directory.Exists(ImageFolderPath))
                                        System.IO.Directory.CreateDirectory(ImageFolderPath);

                                    img.Save(System.IO.Path.Combine(ImageFolderPath, filename));

                                    imageNames.Add(filename);

                                    _summary.NumberOfImagesExported++;

                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}' succeeded.", 2);
                                }
                                catch(Exception ex)
                                {
                                    VerboseMessage($"Exporting image with RelationShipId '{imagedata.RelationshipId}' failed.", 2);

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
            _summary.CouldExportImages = _summary.NumberOfImagesFound > 0 && _summary.NumberOfImagesExported > 0;

            if (System.IO.Directory.Exists(ImageFolderPath))
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

                            if(System.IO.File.Exists(ImageZipPath))
                            {
                                _summary.ImageLocation = ImageZipPath;
                                _summary.CouldZipExportedImages = true;

                                VerboseMessage("Zipping exported images succeeded.", 1);

                                System.IO.Directory.Delete(ImageFolderPath, true);

                                if(!System.IO.Directory.Exists(ImageFolderPath))
                                { 
                                    VerboseMessage($"Deleting Directory '{ImageFolderPath}' succeeded.", 1);
                                }
                                else
                                {
                                    _summary.Errors.Add($"Deleting Directory '{ImageFolderPath}' failed.");

                                    VerboseMessage($"Deleting Directory '{ImageFolderPath}' failed.", 1);
                                }
                            }
                            else
                            {
                                _summary.ImageLocation = ImageFolderPath;
                                _summary.CouldZipExportedImages = false;

                                _summary.Errors.Add("Could not zip exported images.");

                                VerboseMessage($"Zipping exported images failed. Images are still in folder '{ImageFolderPath}'.", 1);
                            }

                            
                        }
                        catch(Exception ex)
                        {
                            _summary.Errors.Add($"Could not create Zip-File because of '{ex.Message}'.");

                            VerboseMessage("Zipping exported images failed.", 1);
                        }
                    }
                    else
                    {
                        _summary.ImageLocation = ImageFolderPath;
                    }
                }
                else
                {
                    VerboseMessage("No images were exported but image output directory exists. Deleting directory.", 1);
                    System.IO.Directory.Delete(ImageFolderPath);
                }
            }
        }

        private void VerboseMessage(string message, int level = 0)
        {
            var indent = level > 0 ?  "\t".Repeat(level) : string.Empty;

            if (_options.Verbose)
                Console.WriteLine(indent + message);
        }

        private void InformationMessage(string message, int level = 0)
        {
            var indent = level > 0 ? "\t".Repeat(level) : string.Empty;

            Console.WriteLine(indent + message);
        }
    }
}
