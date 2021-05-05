using CommandLine;
using CommandLine.Text;
using ConsoleTables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace WordTableExtractor
{
    class Program
    {
        static readonly string ProductVersion = typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion;
        
        static void Main(string[] args)
        {
            var headingInfo = new HeadingInfo("Word Table Extractor", ProductVersion);
            
            Console.WriteLine(headingInfo);
            Console.WriteLine(CopyrightInfo.Default);
            Console.WriteLine();


            var parser = new CommandLine.Parser(with => with.HelpWriter = null);

            var parserResult = parser.ParseArguments<ExportOptions>(args);

            parserResult
                .WithParsed<ExportOptions>(HandleExtraction)
                .WithNotParsed(errs => DisplayHelp(errs, parserResult));
        }

        static void DisplayHelp<T>(IEnumerable<Error> errs, ParserResult<T> result)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = string.Empty;
                h.Copyright = string.Empty;

                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);
        }

        private static void HandleExtraction(ExportOptions options)
        {
            var extractor = new TableExtractor(options);
            
            var summary = extractor.ExtractTables();

            DisplaySummary(summary);
        }

        private static void DisplaySummary(Summary summary)
        {
            Console.WriteLine();
            Console.WriteLine("SUMMARY");
            Console.WriteLine();

            var table = new ConsoleTable("Indicator", "Value");

            table.AddRow("Input file", summary.InputFilePath);
            table.AddRow("Could read input file", summary.CouldReadWordDocument);

            table.AddRow("Output file", summary.OutputFilePath);
            table.AddRow("Could write output file", summary.CouldWriteExcelDocument);

            table.AddRow("Should export images", summary.ShouldExportImages);
            table.AddRow("Could export images", summary.CouldExportImages);

            table.AddRow("Should zip images", summary.ShouldZipExportedImages);
            table.AddRow("Could zip images", summary.CouldZipExportedImages);

            table.AddRow("Tables found", summary.NumberOfTablesFound);
            table.AddRow("Tables imported", summary.NumberOfTablesImported);
            table.AddRow("Tables exported", summary.NumberOfTablesExported);

            table.AddRow("Images found", summary.NumberOfImagesFound);
            table.AddRow("Images exported", summary.NumberOfImagesExported);
            table.AddRow("Image location", summary.ImageLocation);

            table.Write(Format.Alternative);

            Console.WriteLine("ERRORS");
            Console.WriteLine();

            if (summary.Errors.Count > 0)
            {
                var errorTable = new ConsoleTable("#", "Error");

                for(int i = 0; i < summary.Errors.Count; i++)
                {
                    errorTable.AddRow(i + 1, summary.Errors[i]);
                }

                errorTable.Write();
            }
            else
            {
                Console.WriteLine("No errors happened.");
            }
        }
    }
}
