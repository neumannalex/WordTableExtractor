using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Core;

namespace WordTableExtractor
{
    [Verb("test", HelpText = "Testing")]
    public class TestOptions : IOptions
    {
        public string Filename { get; set; }
        public string Output { get; set; }

        [Option('t', "text", Required = true, HelpText = "Some text")]
        public string Text { get; set; }

        [Option('f', "flag", HelpText = "A flag that is true if the value is set, otherwise it is false.")]
        public bool Flag { get; set; }

        [Option('s', "switch", Required = false, Default = true, HelpText = "A switch that is true by default.")]
        public bool? Switch { get; set; }
    }
}
