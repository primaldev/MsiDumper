using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;
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
            parseShortcuts();


            Console.Write(genXML(msiProperties));
        }


        private string genXML(MsiProperties msiProperties)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(MsiProperties));
            using (StringWriter writer = new StringWriter())
            {
                serializer.Serialize(writer, msiProperties);

                return writer.ToString();
            }

        }

        private void parseShortcuts()
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Shortcut`");

            Record record = view.Fetch();

            while (record != null)
            {
                MsiShortCuts msiShortcut = new MsiShortCuts();
                msiShortcut.ShortCut = record.get_StringData(1);
                msiShortcut.StartMenuDirectory = resolveMsiVar(record.get_StringData(2));
                msiShortcut.Name = record.get_StringData(3);
                msiShortcut.Component = record.get_StringData(4);

                if (isAdvertised(record.get_StringData(5))) 
                {
                    MsiComponent msiComponent = getComponent(record.get_StringData(5));

                    if (msiComponent != null)
                    {
                       msiShortcut.ShortCutTarget = resolveMsiVar(msiComponent.KeyPath);
                    }
                    else
                    {
                        Console.WriteLine("Error getting shortcut component");
                    }

                }
                else
                {
                   msiShortcut.ShortCutTarget = resolveMsiVar(record.get_StringData(5));
                }


                msiShortcut.workingDir = record.get_StringData(12);
                //msiProperties.Shortcuts.Add(msiShortcut);
                record = view.Fetch();
            }
        }

   
        //enumerate the Directory table to build a path string
        private string getFullDirectoryPath(string directory)
        {
            
            String fullDirPath = "";
            Boolean hasParent = true;
            int loopcount = 256; //prevent endless loop

            if (directory != null)
            {
                MsiDirectory msiDirectory = (MsiDirectory)directoryTable[directory];

                while (hasParent)
                {
                    
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


        private string resolveMsiVar(String varNames) //
        {
           
            string fullPath = "";
            string[] varNameAr = varNames.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string varName in varNameAr)
            {
                
                    if (varName[0].Equals("#") || varName[0].Equals("!"))
                    {
                        string cvar = Regex.Replace(varName, @"[\#\!]+", "");
                        MsiFile msiFile = getFile(cvar);
                        if (msiFile != null)
                        {

                            fullPath += "\\" + getComponentFullDir(msiFile.Component);                     
                            
                        }

                    }
                    else if (varName[0].Equals("$"))
                    {
                        string evar = Regex.Replace(varName, @"[\#\!]+", "");
                        fullPath += "\\" + getComponentFullDir(evar);

                    }
                    else
                    {
                        
                        //check custom action table                        
                        string resProp = getCustomActionPathSet(varName); //TODO: resolve the condition field as custom action might not be used

                        if (resProp != null)
                        {
                            if (resProp.IndexOf("[") > 0)
                            {
                                fullPath += "\\" + resolveMsiVar(resProp);
                            }
                            else
                            {
                                fullPath += "\\" + getFullDirectoryPath(resProp);
                            }

                            return fullPath;
                        }

                        string rsProp = getMsiProperty(varName);

                        if (rsProp != null)
                        {
                            fullPath += "\\" + rsProp;
                        }

                        fullPath += "\\" + getFullDirectoryPath(getCustomActionPathSet(varName));

                    }
                
            }
            return fullPath;
        }

        private string getComponentFullDir(string component) //limited
        {
            Console.WriteLine("GetcomponentFulldir");
            MsiComponent msiComponent = getComponent(component);
            if (msiComponent != null)
            {
                return getFullDirectoryPath(msiComponent.Directory);
            }

            return ""; //not null but emty string
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
            if (record != null)
            {
                msiComponent.ComponentName = record.get_StringData(1);
                msiComponent.ComponentId = record.get_StringData(2);
                msiComponent.Directory = record.get_StringData(3);
                msiComponent.Atribute = record.get_IntegerData(4);
                msiComponent.Condition = record.get_StringData(5);
                msiComponent.KeyPath = record.get_StringData(6);
                
            }


            return msiComponent;
        }

        private string getMsiProperty(string property)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Property` where Property='" + property + "'");

            Record record = view.Fetch();
            if (record != null)
            {
                return record.get_StringData(2);
            }

            return null;
        }

        private MsiFile getFile(string File)
        {
             WindowsInstaller.View view = queryMsi("SELECT * FROM `File` where File='" + File + "'");

            Record record = view.Fetch();
            MsiFile msiFile = new MsiFile();
            if (record != null)
            {
                msiFile.File = record.get_StringData(1);
                msiFile.Component = record.get_StringData(2);
                msiFile.FileName = record.get_StringData(3);
                msiFile.FileSize = record.get_IntegerData(4);
                msiFile.Version = record.get_StringData(5);
                msiFile.Language = record.get_StringData(6);
                msiFile.Attribute = record.get_IntegerData(7);
                
                
            }

            return msiFile;

        }

        private MsiRegistry getRegistry(string Registry)
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Registry` where Registry='" + Registry + "'");

            Record record = view.Fetch();
            MsiRegistry msiRegistry = new MsiRegistry();
            if (record != null)
            {
                msiRegistry.Registry = record.get_StringData(1);
                msiRegistry.Root = record.get_IntegerData(2);
                msiRegistry.Key = record.get_StringData(3);
                msiRegistry.Name = record.get_StringData(4);
                msiRegistry.Value = record.get_StringData(5);
                msiRegistry.Component = record.get_StringData(6);
                
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

        private string getCustomActionPathSet(String varName) //TODO: resolve the condition field as custom action might not be used
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
