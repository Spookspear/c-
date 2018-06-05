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
        public static string strFilename = "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml";

        // public string strFilename = CommonExcelClasses.strFilename;

        // might pout it in here public string strFilename = CommonExcelClasses.strFilename;

        public static void SaveData(object obj, string strFileName)
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

        public static string readProperty(string strWhichProperty)
        {

            string strRetVal = "Could not Find"; ;
            // load data
            if (File.Exists(strFilename))
            {
                XmlSerializer xs = new XmlSerializer(typeof(InformationFromSettingsForm));
                FileStream read = new FileStream(strFilename, FileMode.Open, FileAccess.Read, FileShare.Read);
                InformationFromSettingsForm info = (InformationFromSettingsForm)xs.Deserialize(read);
                if (strWhichProperty == "strCompareOrColour")
                {
                    strRetVal = info.Differences;
                    read.Close();
                    return strRetVal;
                }

            }

            return strRetVal;

        }

    }
}
