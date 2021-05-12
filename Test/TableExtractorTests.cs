using FluentAssertions;
using System;
using WordTableExtractor;
using WordTableExtractor.Features.Extract;
using Xunit;

namespace Test
{
    public class TableExtractorTests
    {
        private ExtractOptions _options = new ExtractOptions
        {
            Filename = "",
            ExportImages = true,
            ZipImages = true,
            Verbose = false
        };

        [Fact]
        public void RunExtraction()
        {
            var extractor = new TableExtractor(_options);

            var summary = extractor.ExtractTables();
        }

        [Fact]
        public void OutputFolderPath_Is_Correct_If_Output_Is_Default()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFolderPath.Should().Be(@"C:\Daten\Unterordner\Datei");
        }

        [Fact]
        public void OutputFolderPath_Is_Correct_If_Output_Is_Custom_With_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = @"C:\Daten\Neuer_Ordner\Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFolderPath.Should().Be(@"C:\Daten\Neuer_Ordner\Output");
        }

        [Fact]
        public void OutputFolderPath_Is_Correct_If_Output_Is_Custom_Without_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = "Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFolderPath.Should().Be("Output");
        }



        [Fact]
        public void OutputFilePath_Is_Correct_If_Output_Is_Default()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFilePath.Should().Be(@"C:\Daten\Unterordner\Datei\Datei.xlsx");
        }

        [Fact]
        public void OutputFilePath_Is_Correct_If_Output_Is_Custom_With_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = @"C:\Daten\Neuer_Ordner\Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFilePath.Should().Be(@"C:\Daten\Neuer_Ordner\Output\Output.xlsx");
        }

        [Fact]
        public void OutputFilePath_Is_Correct_If_Output_Is_Custom_Without_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = "Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.OutputFilePath.Should().Be(@"Output\Output.xlsx");
        }



        [Fact]
        public void ImageFolderPath_Is_Correct_If_Output_Is_Default()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageFolderPath.Should().Be(@"C:\Daten\Unterordner\Datei\images");
        }

        [Fact]
        public void ImageFolderPath_Is_Correct_If_Output_Is_Custom_With_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = @"C:\Daten\Neuer_Ordner\Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageFolderPath.Should().Be(@"C:\Daten\Neuer_Ordner\Output\images");
        }

        [Fact]
        public void ImageFolderPath_Is_Correct_If_Output_Is_Custom_Without_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = "Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageFolderPath.Should().Be(@"Output\images");
        }



        [Fact]
        public void ImageZipPath_Is_Correct_If_Output_Is_Default()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageZipPath.Should().Be(@"C:\Daten\Unterordner\Datei\images.zip");
        }

        [Fact]
        public void ImageZipPath_Is_Correct_If_Output_Is_Custom_With_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = @"C:\Daten\Neuer_Ordner\Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageZipPath.Should().Be(@"C:\Daten\Neuer_Ordner\Output\images.zip");
        }

        [Fact]
        public void ImageZipPath_Is_Correct_If_Output_Is_Custom_Without_Path()
        {
            var options = new ExtractOptions
            {
                Filename = @"C:\Daten\Unterordner\Datei.docx",
                Output = "Output.xlsx",
                ExportImages = true,
                ZipImages = true
            };

            var extractor = new TableExtractor(options);

            extractor.ImageZipPath.Should().Be(@"Output\images.zip");
        }
    }
}
