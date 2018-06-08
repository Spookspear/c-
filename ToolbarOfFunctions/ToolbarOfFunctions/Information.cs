using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolbarOfFunctions
{
    public class InformationForSettingsForm
    {

        public bool LargeButtons { get; set; }

        public bool HideText { get; set; }

        public string CompareOrColour { get; set; }
        public string CompareOrColourNew { get; set; }

        public string HighLightOrDelete { get; set; }

        public string HighLightOrDeleteNew { get; set; }

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

        public string ColourBack_Found { get; set; }

        public string ColourFore_NotFound { get; set; }

        public string ColourBack_NotFound { get; set; }

        public decimal TimeSheetRowNo { get; set; }

        public bool TimeSheetGetRowNo { get; set; }

        public decimal PingSheetRowNo { get; set; }

        public decimal ColPingRead { get; set; }

        public decimal ColPingWrite { get; set; }

        public bool TestCode { get; set; }

        

    }

}
