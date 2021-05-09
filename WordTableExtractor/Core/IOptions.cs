using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Core
{
    public interface IOptions
    {
        [Option('f', "filename", Required = true, HelpText = "Input filename.")]
        public string Filename { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output filename.")]
        public string Output { get; set; }
    }
}
