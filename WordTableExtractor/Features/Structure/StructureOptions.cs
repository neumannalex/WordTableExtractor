using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor.Features.Structure
{
    [Verb("structure", HelpText = "Analyzes and shows the hierachial structure of data in an Excel table.")]
    public class StructureOptions : IOptions
    {
        public string Filename { get; set; }

        public string Output { get; set; }

        [Option('s', "sheet", Required = true, HelpText = "Worksheet that contains the data.")]
        public string Sheet { get; set; }

        [Option('r', "range", Required = true, HelpText = "Data range in input file (e. g. A2:C10).")]
        public string Range { get; set; }
    }
}
