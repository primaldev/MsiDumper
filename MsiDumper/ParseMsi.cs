using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsInstaller;

namespace MsiDumper
{
    class ParseMsi
    {
       
        public ParseMsi(string msiFile)
        {
            Database database = getInstaller().OpenDatabase(msiFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
            parseDatabase(database);
        }

        public ParseMsi(string msiFile, string transForms)
        {
            Database database = getInstaller().OpenDatabase(msiFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);
            foreach (String transForm in transForms.Split(new string[] { "," }, StringSplitOptions.None))
            {
                database.ApplyTransform(transForm, MsiTransformError.msiTransformErrorNone);
            }

            parseDatabase(database);
        }


        private Installer getInstaller()
        {
            Type type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            return (Installer)Activator.CreateInstance(type);
        }


        private void parseDatabase(Database database)
        {
            WindowsInstaller.View view = null;
            try
            {
                view = database.OpenView("SELECT * FROM `Property`");
                view.Execute();
            }
            catch (Exception ex)
            {
                
                Console.WriteLine("I made a boo boo, " + ex.Message);
            }
            

            Record properties = view.Fetch();
            while (properties !=null)
            {
                Console.WriteLine(properties.get_StringData(1));
                properties = view.Fetch();
            }

           

        }




        
    }
}
