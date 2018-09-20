using System;
using System.Collections.Generic;

namespace RiggingConsoleApp
{
    public class RiggingHeaderDS
    {
        public int Id { get; set; }

        public string FileName { get; set; }

        public DateTime FileDate { get; set; }

        public string ContactPerson { get; set; }

        public string BudgetHolder { get; set; }

        public string VesselLocation { get; set; }

        public string ProjectDepartment { get; set; }

        public string DateRequested { get; set; }

        public string DateRequired { get; set; }

        public string ProjectDuration { get; set; }

        public string SAPCostCode { get; set; }

        public string DeliveryDetails { get; set; }

        public string Remarks { get; set; }

        public string ATRWONO { get; set; }

        public string Vendor { get; set; }

        public string PONumber { get; set; }

        public List<RiggingLinesDS> lstRiggingLines { get; set; }

    }

    public class RiggingLinesDS
    {
        public int Id { get; set; }

        public string HighLevelDesc { get; set; }

        public string LowLevelDesc { get; set; }

        public string Quantity { get; set; }

        public string ItemValue { get; set; }

        public string TotalValue { get; set; }

        public string TestProcedure { get; set; }

        public string LineOrAdditional { get; set; }

        // public Guid LinkToHeader { get; set; }

        public int RiggingHeaderId { get; set; }

        public virtual RiggingHeaderDS RiggingHeader { get; set; }
    }
}