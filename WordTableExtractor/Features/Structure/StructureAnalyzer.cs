using System;
using System.Collections.Generic;
using System.Text;
using WordTableExtractor.Services;

namespace WordTableExtractor.Features.Structure
{
    public class StructureAnalyzer
    {
        private readonly StructureOptions _options;

        public StructureAnalyzer(StructureOptions options)
        {
            _options = options;
        }

        public void ViusalizeStructure()
        {
            var importer = new ExcelTableImporter();

            var tree = importer.ImportAsTree(_options.Filename, _options.Sheet, _options.Range, _options.HierarchyColumn, _options.TypeColumn, _options.DirectoryType, _options.Captions);            

            var dump = tree.TreeRoot.ToText("", ExcelTableImporter.FormatTreenodeForDump);

            if(string.IsNullOrEmpty(_options.Output))
            {
                Console.WriteLine("Structure of imported data\n:");

                Console.WriteLine(dump);
            }
            else
            {
                System.IO.File.WriteAllText(_options.Output, dump);
                Console.WriteLine($"Structure of imported data was written to '{_options.Output}'.");
            }
        }
    }
}
