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

            var helpWriter = new StringWriter();

            var parser = new CommandLine.Parser(with => with.HelpWriter = helpWriter);

            parser.ParseArguments<Options>(args)
                .WithParsed(TestHandler)
                .WithNotParsed(errs => DisplayHelp(errs, helpWriter));

            Console.WriteLine($"Exit code= {Environment.ExitCode}");
        }

        static void DisplayHelp(IEnumerable<Error> errs, TextWriter helpWriter)
        {
            if (errs.IsVersion() || errs.IsHelp())
                Console.WriteLine(helpWriter.ToString());
            else
                Console.Error.WriteLine(helpWriter.ToString());
        }

        private static void HandleExtraction(Options options)
        {
            var extractor = new TableExtractor(options);
            extractor.ExtractTables();
        }

        private static void TestHandler(Options options)
        {
            options.Dump();
        }
    }
}
