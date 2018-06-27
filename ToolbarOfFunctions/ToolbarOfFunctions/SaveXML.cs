#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml.Serialization;
using System.IO;

using ToolbarOfFunctions_CommonClasses;

namespace ToolbarOfFunctions
{
    public class SaveXML
    {

        public static string strFileName = "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml";

        // public string strFileName = CommonExcelClasses.strFileName;

        // might pout it in here public string strFileName = CommonExcelClasses.strFileName;

            /*
        public static void SaveData(object obj)
        {
            try
            {
                XmlSerializer sr = new XmlSerializer(obj.GetType());
                TextWriter writer = new StreamWriter(strFileName);
                sr.Serialize(writer, obj);

                writer.Close();
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

            */


        /*
        public static string readProperty(string strWhichProperty)
        {

            string strRetVal = "Could not Find"; ;
            // load data
            if (File.Exists(strFileName))
            {
                XmlSerializer xS = new XmlSerializer(typeof(InformationForSettingsForm));
                FileStream fsRead = new FileStream(strFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                InformationForSettingsForm info = (InformationForSettingsForm)xS.Deserialize(fsRead);
                if (strWhichProperty == "strCompareOrColour")
                {
                    strRetVal = info.CompareOrColour;
                    fsRead.Close();
                    return strRetVal;
                }

            }

            return strRetVal;

        }

        */

        /*

        // purpose of this is to load the data into the InformationForSettingsForm class
        public static InformationForSettingsForm LoadData()
        {
            XmlSerializer xS = new XmlSerializer(typeof(InformationForSettingsForm));
            FileStream fsRead = new FileStream(strFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            InformationForSettingsForm infoLocal = (InformationForSettingsForm)xS.Deserialize(fsRead);

            fsRead.Close();

            return infoLocal;

        }

    */



    }

}
