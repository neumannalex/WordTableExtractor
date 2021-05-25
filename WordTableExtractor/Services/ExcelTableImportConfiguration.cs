using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Services
{
    public class ExcelTableImportConfiguration
    {
        public string Filename { get; set; }
        public string Sheet { get; set; }
        public string Range { get; set; }
        public bool HasHeader { get; set; }
        public string RowTypeColumn { get; set; }
        public IEnumerable<string> FolderRowTypes { get; set; } = new List<string>();
        public IEnumerable<string> LeafRowTypes { get; set; } = new List<string>();
    }
}
