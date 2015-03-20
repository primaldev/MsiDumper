using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsInstaller;

namespace MsiDumper
{
    class Program

    {
        
        static void Main(string[] args)
        {
            new ParseMsi("d:\\temp\\Configuration_Manager_2012_R2_SDK_Setup.msi");

        }


       
    }
}
