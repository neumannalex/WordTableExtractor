using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor.Services
{
    public class ExcelTableImportOptions : IOptions
    {
        public string Filename { get; set; }

        public string Output { get; set; }

        [Option('s', "sheet", Required = true, HelpText = "Worksheet that contains the data.")]
        public string Sheet { get; set; }

        [Option('r', "range", Required = true, HelpText = "Data range in input file (e. g. A2:C10).")]
        public string Range { get; set; }

        [Option('c', "captions", Required = false, HelpText = "Set to true if the first row of the range is a header.")]
        public bool Captions { get; set; }

        [Option('h', "hierarchy", Required = true, HelpText = "Name of the column that contains hierarchy info like '2.1.3'.")]
        public string HierarchyColumn { get; set; }

        [Option('t', "type", Required = true, HelpText = "Name of the column that contains hierarchy info like '2.1.3'.")]
        public string TypeColumn { get; set; }

        [Option('d', "directory", Required = true, HelpText = "Values in the Type column that represent directory-like nodes.")]
        public string DirectoryType { get; set; }
    }
}
