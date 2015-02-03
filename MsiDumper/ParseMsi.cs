using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        //go through the levels of msi to determen the final value of a property
        private string resolveProperty(string property)
        {

            string resProp;

            

            resProp = getCustomActionPathSet(property);

            if (resProp != null)
            {
                return resProp; // not yet
            }

            resProp = getMsiProperty(property);

            if (resProp != null)
            {
                return resProp;
            }



            //1) custom action
            //2) property            
            //3) directory table

            return null;
        }

        //enumerate the Directory table to build a path string
        private string getFullDirectoryPath(string directory)
        {
            String fullDirPath = "";
            Boolean hasParent = true;
            int loopcount = 256; //prevent endless loop

            MsiDirectory msiDirectory = (MsiDirectory) directoryTable[directory];
            while(hasParent){
                
                if (msiDirectory != null) 
                {
                    fullDirPath = fullDirPath + "\\" + msiDirectory.getDefaultDirLongName();


                    if (msiDirectory.DirectoryParent != null & msiDirectory.DirectoryParent.Length > 0 && msiDirectory.Directory.ToLower().Equals(msiDirectory.DirectoryParent.ToLower()) && loopcount > 0)
                    {
                        msiDirectory = (MsiDirectory)directoryTable[directory];
                    }
                    else
                    {
                        hasParent = false;
                    }

                    loopcount--;
                }
               
            }
            return fullDirPath;

        }


        private string resolveMsiVar(String varName) //
        {
            if (varName[0].Equals("[") && varName[varName.Length -1].Equals("]")) 
            {

                if (varName[1].Equals("#") || varName[1].Equals("!")  )
                {
                    string cvar = Regex.Replace(varName, @"[\[\]\#\!]+", "");
                    MsiFile msiFile = getFile(cvar);
                    if (msiFile != null)
                    {
                        MsiComponent msiComponent = getComponent(msiFile.Component);

                        if (msiComponent != null)
                        {
                            return msiComponent.Directory; //not yet
                        }
                    }

                    
                } 
                else if(varName[1].Equals("$"))                 
                {


                }              
                else
                {
                    //check for featuretable
                    //check for properties
                    //Custom action overule all


                }
            }
            return null;
        }


        private bool isAdvertised(string Name)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Feature` where Feature='" + Name + "'");
            Record record = view.Fetch();

            if (record != null)
            {
                return true;
            }

            return false;
        }

        private MsiComponent getComponent(string componentName)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Component` where Component='" + componentName + "'");

            Record record = view.Fetch();
            MsiComponent msiComponent = new MsiComponent();
            while (record != null)
            {
                msiComponent.ComponentName = record.get_StringData(1);
                msiComponent.ComponentId = record.get_StringData(2);
                msiComponent.Directory = record.get_StringData(3);
                msiComponent.Atribute = record.get_IntegerData(4);
                msiComponent.Condition = record.get_StringData(5);
                msiComponent.KeyPath = record.get_StringData(6);
                view.Fetch();
            }


            return msiComponent;
        }

        private string getMsiProperty(string property)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Property` where Property='" + property + "'");

            Record record = view.Fetch();

            return record.get_StringData(2);
        }

        private MsiFile getFile(string File)
        {
             WindowsInstaller.View view = queryMsi("SELECT * FROM `File` where File='" + File + "'");

            Record record = view.Fetch();
            MsiFile msiFile = new MsiFile();
            while (record != null)
            {
                msiFile.File = record.get_StringData(1);
                msiFile.Component = record.get_StringData(2);
                msiFile.FileName = record.get_StringData(3);
                msiFile.FileSize = record.get_IntegerData(4);
                msiFile.Version = record.get_StringData(5);
                msiFile.Language = record.get_StringData(6);
                msiFile.Attribute = record.get_IntegerData(7);
                
                view.Fetch();
            }

            return msiFile;

        }

        private MsiRegistry getRegistry(string Registry)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Registry` where Registry='" + Registry + "'");

            Record record = view.Fetch();
            MsiRegistry msiRegistry = new MsiRegistry();
            while (record != null)
            {
                msiRegistry.Registry = record.get_StringData(1);
                msiRegistry.Root = record.get_IntegerData(2);
                msiRegistry.Key = record.get_StringData(3);
                msiRegistry.Name = record.get_StringData(4);
                msiRegistry.Value = record.get_StringData(5);
                msiRegistry.Component = record.get_StringData(6);
                view.Fetch();
            }

            return msiRegistry;
        }


        private WindowsInstaller.View queryMsi(String query)
        {
            WindowsInstaller.View view = null;
            try
            {
                view = database.OpenView(query);
                view.Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to query msi, " + ex.Message);
            }


            return view;
        }

        private string getCustomActionPathSet(String varName)
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
