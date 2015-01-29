using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiProperties
    {
        public string Name { get; set; }
        public List<MsiShortCuts> Shortcuts { get; set; }

    }
}
