using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsiDumper
{
    class MsiProperties
    {


        public string ArpComments { get; set; }
        public string Manufacturer { get; set; }
        public string ProductName { get; set; }
        public string ArpComments { get; set; }
        public string ArpHelpLink { get; set; }
        public string ArpContact { get; set; }
        public string AllUsers { get; set; }
        public string ArpInstallLocation { get; set; }
        public string ArpTelephone { get; set; }
        public string PrimaryFolder { get; set; }
        public string MsiInstallPerUser { get; set; }
        public string Reboot { get; set; }
        public string RootDrive { get; set; }
        public string TargetDir { get; set; }
        public string InstallDir { get; set; }
        public string UpgradeCode { get; set; }
        public string ProductCode { get; set; }
        public string ProductVersion { get; set; }
        public string ProductLanguage { get; set; }

        public List<MsiShortCuts> Shortcuts { get; set; }
        public List<MsiCustomAction> customActionsModified { get; set; }

        //custom properties from transform
        //added directories from transform


    }
}

