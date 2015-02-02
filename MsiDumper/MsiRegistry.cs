using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiRegistry
    {
        public string Registry { get; set; }
        public int Root { get; set; }
        public string Key { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }
        public string Component { get; set; }
    }
}
