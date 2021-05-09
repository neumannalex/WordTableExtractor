using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor.Import
{
    [Verb("transform", HelpText = "Transform excel tables.")]
    public class ImportOptions : IOptions
    {
        public string Filename { get; set; }

        public string Output { get; set; }

        [Option('s', "sheet", Required = true, HelpText = "Worksheet that contains the data.")]
        public string Sheet { get; set; }

        [Option('r', "range", Required = true, HelpText = "Data range in input file (e. g. A2:C10).")]
        public string Range { get; set; }

        [Option('c', "consistency", Required = false, HelpText = "Only checks consistency of input hierarchy.")]
        public bool Consistency { get; set; }
    }
}
