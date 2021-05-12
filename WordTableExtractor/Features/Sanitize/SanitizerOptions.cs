using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor.Features.Sanitize
{
    [Verb("sanitize", HelpText = "Sanitizes the numbering of hierachial rows in an Excel sheet.")]
    public class SanitizerOptions : IOptions
    {
        public string Filename { get; set; }

        public string Output { get; set; }

        [Option('s', "sheet", Required = true, HelpText = "Worksheet that contains the data.")]
        public string Sheet { get; set; }

        [Option('r', "range", Required = true, HelpText = "Data range in input file (e. g. A2:C10).")]
        public string Range { get; set; }

        [Option('c', "column", Required = true, HelpText = "Name of the column header that contains the numbering (e. g. 'section').")]
        public string Column { get; set; }
    }
}
