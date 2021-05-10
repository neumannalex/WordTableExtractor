using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor.Export
{
    [Verb("export", HelpText = "Exports excel tables into a template.")]
    public class ExportOptions : IOptions
    {
        public string Filename { get; set; }

        public string Output { get; set; }

        [Option('s', "sheet", Required = true, HelpText = "Worksheet that contains the data.")]
        public string Sheet { get; set; }

        [Option('r', "range", Required = true, HelpText = "Data range in input file (e. g. A2:C10).")]
        public string Range { get; set; }

        [Option('t', "template", Required = true, HelpText = "Worksheet that acts as a template.")]
        public string Template { get; set; }
    }
}
