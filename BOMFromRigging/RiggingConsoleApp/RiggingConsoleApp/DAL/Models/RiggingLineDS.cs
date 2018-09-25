using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiggingConsoleApp.DAL.Models
{
    public class RiggingLineDS
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

        public virtual RiggingConsoleApp.DAL.Models.RiggingHeaderDS RiggingHeader { get; set; }
    }
}
