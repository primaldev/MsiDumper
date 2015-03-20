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
        Session msiSession;

        public ParseMsi(string msiFile)
        {
            try
            {
                this.database = getInstaller().OpenDatabase(msiFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
                
            }
            catch (Exception ex)
            {
                Console.Write("Error opening msi database. " + ex.Message);
                Environment.ExitCode = 1;

            }
            if (database != null)
            {
                parseDatabase();
            }
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
            getInstaller().UILevel = MsiUILevel.msiUILevelNone;
           msiSession = getInstaller().OpenPackage(database, 0);
           
           msiSession.DoAction("CostInitialize");
           msiSession.DoAction("CostFinalize"); 

            msiProperties = new MsiProperties();
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

        private string cleanShortCutName(string shortCutName)
        {                       
           return Path.GetFileNameWithoutExtension(shortCutName.Split(new Char[] { '|' })[1]);          
        }


        private string getTargetPath(string target)
        {

            try
            {
                return msiSession.get_TargetPath(target);
            }
            catch (Exception)
            {
                
               //ignore
            }

            return null;
        
        
        }



        private void parseShortcuts()
        {
            WindowsInstaller.View view = queryMsi("SELECT * FROM `Shortcut`");

            Record record = view.Fetch();
            List<MsiShortCuts> shortCuts = new List<MsiShortCuts>();
            while (record != null)
            {
                MsiShortCuts msiShortcut = new MsiShortCuts();
                msiShortcut.ShortCut = record.get_StringData(1);
                msiShortcut.StartMenuDirectory = resolveMsiVar(record.get_StringData(2));
                msiShortcut.Name = cleanShortCutName(record.get_StringData(3));
                msiShortcut.Component = record.get_StringData(4);

                if (isAdvertised(record.get_StringData(5))) 
                {
                    Console.WriteLine("ishAdvertised...");
                    MsiComponent msiComponent = getComponent(record.get_StringData(4));

                    if (msiComponent != null)
                    {
                        msiShortcut.ShortCutTarget = getTargetPath(msiComponent.Directory) + getFile(msiComponent.KeyPath).FileName;
                        Console.WriteLine("ishAdvertised with keypath..." + msiComponent.KeyPath + "::" + getTargetPath(msiComponent.Directory)  + getFile(msiComponent.KeyPath).FileName);
                        
                    }
                    else
                    {                        
                        Environment.ExitCode = 2;
                    }

                }
                else
                {
                    msiShortcut.ShortCutTarget = resolveMsiVar(record.get_StringData(5));
                }


                msiShortcut.workingDir = resolveMsiVar(record.get_StringData(12));
                shortCuts.Add(msiShortcut);
                record = view.Fetch();
            }

            msiProperties.Shortcuts = shortCuts;


        }

        
   
        //enumerate the Directory table to build a path string


        private string resolveMsiVar(String varNames) //
        {
           
            string fullPath = "";
            if (varNames != null)
            {
                string[] varNameAr = varNames.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string varName in varNameAr)
                {

                    if (varName[0].Equals("#") || varName[0].Equals("!"))
                    {
                        string cvar = Regex.Replace(varName, @"[\#\!]+", "");
                        MsiFile msiFile = getFile(cvar);
                        if (msiFile != null)
                        {

                            fullPath += getTargetPath(msiFile.Component);

                        }

                    }
                    else if (varName[0].Equals("$"))
                    {
                        string evar = Regex.Replace(varName, @"[\#\!]+", "");
                        fullPath +=  getTargetPath(evar);

                    }
                    else
                    {

                        //check custom action table                        
                        string resProp = getCustomActionPathSet(varName); //TODO: resolve the condition field as custom action might not be used

                        if (resProp != null)
                        {
                            if (resProp.IndexOf("]") > 0)
                            {
                                fullPath += resolveMsiVar(resProp);

                            }
                            else
                            {
                                fullPath +=  getTargetPath(resProp);
                            }


                        }
                        else
                        {

                            string rsProp = getMsiProperty(varName);

                            if (rsProp != null)
                            {
                                fullPath +=  rsProp;
                            }
                        }



                    }

                }
                
                return fullPath;
            }
            else
            {
                return null;
            }
            
        }

        private string getComponentFullDir(string component) //limited
        {
            Console.WriteLine("GetcomponentFulldir");
            MsiComponent msiComponent = getComponent(component);
            if (msiComponent != null)
            {
                return getTargetPath(msiComponent.Directory);
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


            try
            {
                return msiSession.get_Property(property);
            }
            catch (Exception)
            {
                
                //ignore
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
            Console.WriteLine("51 " + varName + "--" + resPath);
            return resPath;
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
                directoryTable.Add(properties.get_StringData(1).ToLower(), new MsiDirectory(properties.get_StringData(1), properties.get_StringData(2), properties.get_StringData(3)));
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
                        msiProperties.ProductName = properties.get_StringData(2);
                        break;
                    case "arpcomments":
                        msiProperties.ArpComments = properties.get_StringData(2);
                        break;
                    case "arphelplink":
                        msiProperties.ArpHelpLink = properties.get_StringData(2);
                        break;
                    case "arpcontact":
                        msiProperties.ArpContact = properties.get_StringData(2);
                        break;
                    case "allusers":
                        msiProperties.AllUsers = properties.get_StringData(2);
                        break;
                    case "arpinstalllocation":
                        msiProperties.ArpInstallLocation = properties.get_StringData(2);
                        break;
                    case "arptelephone":
                        msiProperties.ArpTelephone = properties.get_StringData(2);
                        break;
                    case "primaryfolder":
                        msiProperties.PrimaryFolder = properties.get_StringData(2);
                        break;
                    case "msiinstallperuser":
                        msiProperties.MsiInstallPerUser = properties.get_StringData(2);
                        break;
                    case "reboot":
                        msiProperties.Reboot = properties.get_StringData(2);
                        break;
                    case "rootdrive":
                        msiProperties.RootDrive = properties.get_StringData(2);
                        break;
                    case "targetdir":
                        msiProperties.TargetDir = properties.get_StringData(2);
                        break;
                    case "installdir":
                        msiProperties.InstallDir = properties.get_StringData(2);
                        break;
                    case "upgradecode":
                        msiProperties.UpgradeCode = properties.get_StringData(2);
                        break;
                    case "productcode":
                        msiProperties.ProductCode = properties.get_StringData(2);
                        break;
                    case "productversion":
                        msiProperties.ProductVersion = properties.get_StringData(2);
                        break;
                    case "productlanguage":
                        msiProperties.ProductLanguage = properties.get_StringData(2);
                        break;
                }

                properties = view.Fetch();
            }

           

        }



      
    }
}
