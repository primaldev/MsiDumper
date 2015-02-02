using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiFile
    {
        public string File { get; set; }
        public string Component { get; set; }
        public string FileName { get; set; }
        public int FileSize { get; set; }
        public string Version { get; set; }
        public string Language { get; set; }
        public int Attribute { get; set; }

    }
}
