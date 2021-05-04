using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor
{
    public class Options
    {
        [Option('f', "filename", Required = true, HelpText = "Input filename of Word document.")]
        public string Filename { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output filename of Excel document.")]
        public string Output { get; set; }

        [Option('i', "images", Required = false, Default = true, HelpText = "Set to 'true' to export images from tables.")]
        public bool ExportImages { get; set; }

        [Option('z', "zip", Required = false, Default = false, HelpText = "Set to 'true' to put exported images into a zip-archive.")]
        public bool ZipImages { get; set; }

        [Option('v', "verbose", Required = false, Default = false, HelpText = "Set to 'true' to view details of work.")]
        public bool Verbose { get; set; }

        [Usage(ApplicationAlias = "wte")]
        public static IEnumerable<Example> Examples
        {
            get
            {
                return new List<Example>
                {
                    new Example(@"Export tables from C:\Data\Document.docx to C:\Data\Document\Document.xlsx and export images to folder C:\Data\Document\images", new Options{ Filename = @"C:\Data\Document.docx" }),
                    new Example(@"Export tables from C:\Data\Document.docx to C:\Data\Document\Document.xlsx and export images to archive C:\Data\Document\images.zip", new Options{ Filename = @"C:\Data\Document.docx", ZipImages = true }),
                    new Example(@"Export tables from C:\Data\Document.docx to C:\Other\MyFile\MyFile.xlsx and export images to folder C:\Other\MyFile\images", new Options{ Filename = @"C:\Data\Document.docx", Output = @"C:\Other\MyFile.xlsx" }),
                    new Example(@"Export tables from C:\Data\Document.docx to C:\Data\Document\Document.xlsx without exporting images and display verbose messages", new Options{ Filename = @"C:\Data\Document.docx", Verbose = true, ExportImages = false }),
                };
            }
        }

        public void Dump()
        {
            Console.WriteLine($"Filename: {Filename}");
            Console.WriteLine($"Output: {Output}");
            Console.WriteLine($"ExportImages: {ExportImages}");
            Console.WriteLine($"ZipImages: {ZipImages}");
            Console.WriteLine($"Verbose: {Verbose}");
        }
    }
}
