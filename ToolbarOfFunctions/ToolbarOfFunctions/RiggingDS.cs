using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolbarOfFunctions
{
    public class RiggingHeaderDS
    {
        public int MyIndex { get; set; }

        public string FileName { get; set; }

        public DateTime FileDate { get; set; }

        public string ContactPerson { get; set; }

        public string BudgetHolder { get; set; }

        public string VesselLocation { get; set; }

        public string ProjectDepartment { get; set; }

        public string DateRequested { get; set; }

        public string DateRequired { get; set; }

        public string ProjectDuration { get; set; }

        public string SAPCostCode  { get; set; }

        public string DeliveryDetails { get; set; }

        public string Remarks { get; set; }

        public string ATRWONO { get; set; }

        public string Vendor { get; set; }

        public string PONumber { get; set; }

        public List<RiggingLinesDS> lstRiggingLines { get; set; }


    }

    // line items - link to second class?
    public class RiggingLinesDS
    {

        public int MyIndex { get; set; }

        public string HighLevelDesc { get; set; }

        public string LowLevelDesc { get; set; }

        public string Quantity { get; set; }

        public string ItemValue { get; set; }

        public string TotalValue { get; set; }

        public string TestProcedure { get; set; }

        public double GetTotalValue()
        {
            double iItemValue = 0;
            double iQty = 1;

            if (ItemValue.Length > 0)
            {
                iItemValue = Convert.ToDouble(ItemValue);
            }

            if (Quantity.Length > 0)
            {
                iQty = Convert.ToDouble(Quantity);
            }

            return System.Convert.ToDouble(iItemValue * iQty);
        }


    }



}
