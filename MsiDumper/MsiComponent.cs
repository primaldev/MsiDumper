using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiComponent
    {
        public string ComponentName { get; set; }
        public string ComponentId { get; set; }
        public string Directory { get; set; }
        public int Atribute { get; set; }
        public string Condition { get; set; }
        public string KeyPath { get; set; }

    }
}
