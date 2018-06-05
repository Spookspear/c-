using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolbarOfFunctions
{
    public class InformationFromSettingsForm
    {

        public bool LargeButtons { get; set; }

        public bool HideText { get; set; }

        public string Differences { get; set; }

        public string HighLightOrDelete { get; set; }

        public bool DisplayTimeTaken { get; set; }

        public bool ProduceMessageBox { get; set; }

        public string DelModeAorBorC { get; set; }

        public decimal HighlightRowsOver { get; set; }

        public decimal NoOfColumnsToCheck { get; set; }

        public decimal ComparingStartRow { get; set; }

        public decimal DupliateColumnToCheck { get; set; }

        public string ColourFound { get; set; }

        public string ColourNotFound { get; set; }

        public decimal TimeSheetRowNo { get; set; }

        public bool TimeSheetGetRowNo { get; set; }

        public decimal PingSheetRowNo { get; set; }

        public decimal ColPingRead { get; set; }

        public decimal ColPingWrite { get; set; }

    }
}
