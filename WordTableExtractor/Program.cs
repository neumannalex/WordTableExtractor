using CommandLine;
using CommandLine.Text;
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

            var parserResult = parser.ParseArguments<Options>(args);

            parserResult
                .WithParsed(TestHandler)
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

        private static void HandleExtraction(Options options)
        {
            var extractor = new TableExtractor(options);
            
            var summary = extractor.ExtractTables();

            DisplaySummary(summary);
        }

        private static void DisplaySummary(Summary summary)
        {
            Console.WriteLine("This ist the summary.");
        }

        private static void TestHandler(Options options)
        {
            options.Dump();
        }
    }
}
