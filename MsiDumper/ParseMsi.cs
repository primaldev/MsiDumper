using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsInstaller;


namespace MsiDumper
{
    class ParseMsi
    {
        MsiProperties msiProperties;
        Hashtable directoryTable;
        Database database;

        public ParseMsi(string msiFile)
        {
            this.database = getInstaller().OpenDatabase(msiFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
            parseDatabase();
        }

        public ParseMsi(string msiFile, string transForms)
        {
           this.database = getInstaller().OpenDatabase(msiFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
            foreach (String transForm in transForms.Split(new string[] { "," }, StringSplitOptions.None))
            {
                database.ApplyTransform(transForm, MsiTransformError.msiTransformErrorNone);
            }

            parseDatabase();
        }


        private Installer getInstaller()
        {
            Type type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            return (Installer)Activator.CreateInstance(type);
        }


        private void parseDatabase()
        {
            
            parseDirectory();
            parseProperties();
        }

        


        private void getPath(String varName)
        {
            if (varName[0].Equals("[") && varName[varName.Length -1].Equals("]")) 
            {
                if (varName[1].Equals("#") || varName[1].Equals("~") || varName[1].Equals("@"))
                {
                    //parse filname
                }
                else
                {
                    //check for featuretable
                }
            }
        }


        private string getCustomActionPathset(String varName)
        {
            String resPath = null;
            WindowsInstaller.View view = null;
            try
            {
                view = database.OpenView("SELECT * FROM `CustomAction` where Source='" + varName + "'");
                view.Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to query msi, " + ex.Message);
            }
           

            Record properties = view.Fetch();
            while (properties != null)
            {
                int setPropActio = properties.get_IntegerData(2) & 51;

                if (setPropActio == 51)
                {
                    resPath = properties.get_StringData(4);
                }
                properties = view.Fetch();
            }
            return resPath;
        }

        private void parseShortCuts()
        {

        }

        private void parseDirectory()
        {
            WindowsInstaller.View view = null;
            try
            {
                view = database.OpenView("SELECT * FROM `Directory`");
                view.Execute();
            }
            catch (Exception ex)
            {

                Console.WriteLine("Unable to query msi, " + ex.Message);
            }
            directoryTable = new Hashtable();

            Record properties = view.Fetch();
            while (properties != null)
            {
                directoryTable.Add(properties.get_StringData(1), new MsiDirectory(properties.get_StringData(1), properties.get_StringData(2), properties.get_StringData(3)));
                properties = view.Fetch();
            }
        }




        private void parseProperties()
        {
            WindowsInstaller.View view = null;
            try
            {
                view = database.OpenView("SELECT * FROM `Property`");
                view.Execute();
            }
            catch (Exception ex)
            {

                Console.WriteLine("Unable to query msi, " + ex.Message);
            }

            msiProperties = new MsiProperties();
            Record properties = view.Fetch();
            while (properties !=null)
            {
                Console.WriteLine(properties.get_StringData(1));

                switch (properties.get_StringData(1).ToLower())
                {
                    case "manufacturer":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "productname":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "arpcomments":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "arphelplink":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "arpcontact":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "allusers":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "arpinstalllocation":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "arptelephone":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "primaryfolder":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "msiinstallperuser":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "reboot":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "rootdrive":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "targetdir":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "installdir":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "upgradecode":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "productcode":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "productversion":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                    case "productlanguage":
                        msiProperties.Manufacturer = properties.get_StringData(2);
                        break;
                }

                properties = view.Fetch();
            }

           

        }




        
    }
}
