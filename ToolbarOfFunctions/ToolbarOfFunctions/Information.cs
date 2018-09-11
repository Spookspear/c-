using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml.Serialization;
using System.IO;


namespace ToolbarOfFunctions
{
    public class InformationForSettingsForm
    {
        public static string strFileName = "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml";

        public bool LargeButtons { get; set; }

        public bool HideText { get; set; }

        public string CompareOrColour { get; set; }

        public string ColourOrDelete { get; set; }

        public bool DisplayTimeTaken { get; set; }

        public bool ProduceInitialMessageBox { get; set; }

        public bool ProduceCompleteMessageBox { get; set; }

        public string DelModeAorBorC { get; set; }

        public decimal HighlightRowsOver { get; set; }

        public decimal NoOfColumnsToCheck { get; set; }

        public decimal ComparingStartRow { get; set; }

        public decimal DupliateColumnToCheck { get; set; }

        public string ColourFoundText { get; set; }

        public string ColourNotFoundText { get; set; }
        
        public string ColourFore_Found { get; set; }

        public bool ColourBold_Found { get; set; }

        public string ColourBack_Found { get; set; }

        public string ColourFore_NotFound { get; set; }

        public bool ColourBold_NotFound { get; set; }

        public string ColourBack_NotFound { get; set; }

        public decimal TimeSheetRowNo { get; set; }

        public bool TimeSheetGetRowNo { get; set; }

        public decimal PingSheetRowNo { get; set; }

        public decimal ColPingRead { get; set; }

        public decimal ColPingWrite { get; set; }

        public bool TestCode { get; set; }

        public bool RecordTimes { get; set; }

        public bool HideSeperator { get; set; }

        public bool TurnOffScreenValidation { get; set; }

        public bool ClearFormatting { get; set; }

        public string FileDateTime { get; set; }

        public bool ExtractFileName { get; set; }

        public decimal ColExtractedFile { get; set; }

        public decimal ZapStartDefaultRow { get; set; }

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


        public InformationForSettingsForm LoadMyData()
        {
            XmlSerializer xS = new XmlSerializer(typeof(InformationForSettingsForm));
            FileStream fsRead = new FileStream(strFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            InformationForSettingsForm infoLocal = (InformationForSettingsForm)xS.Deserialize(fsRead);

            fsRead.Close();

            return infoLocal;
        }

    }


}
