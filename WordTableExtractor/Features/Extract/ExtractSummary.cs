using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Features.Extract
{
    public class ExtractSummary
    {
        public bool CouldReadWordDocument { get; set; }
        public string InputFilePath { get; set; }

        public int NumberOfTablesFound { get; set; }
        public int NumberOfTablesImported { get; set; }
        public int NumberOfTablesExported { get; set; }

        public int NumberOfImagesFound { get; set; }
        public int NumberOfImagesExported { get; set; }

        
        public bool CouldWriteExcelDocument { get; set; }
        public string OutputFilePath { get; set; }
        
        public bool ShouldExportImages { get; set; }
        public bool CouldExportImages { get; set; }

        public bool ShouldZipExportedImages { get; set; }
        public bool CouldZipExportedImages { get; set; }

        public string ImageLocation { get; set; }

        public List<string> Errors { get; set; } = new List<string>();
    }
}
