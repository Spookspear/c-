using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiggingConsoleApp.DAL.Models
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

        public List<RiggingConsoleApp.DAL.Models.RiggingLineDS> lstRiggingLines { get; set; }


    }
}
