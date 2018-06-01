using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml.Serialization;
using System.IO;

namespace ToolbarOfFunctions
{
    public class SaveXML
    {

        public static void SaveData(object obj, string strFileName)
        {
            XmlSerializer sr = new XmlSerializer(obj.GetType());
            TextWriter writer = new StreamWriter(strFileName);
            sr.Serialize(writer, obj);
            writer.Close();

        }

    }
}
