using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiDirectory
    {

        public string Directory { get; set; }
        public string DirectoryParent { get; set; }
        public string DefaultDir { get; set; }

        public MsiDirectory()
        {
        }
        public MsiDirectory(string Directory, string DirectoryParent, string DefaultDir )
        {
            this.Directory = Directory;
            this.DefaultDir = DefaultDir;
            this.DirectoryParent = DirectoryParent;
            
        }

    }
}
