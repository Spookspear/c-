using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolbarOfFunctions.DAL
{
     public class RiggingContext : DbContext
    {


        public RiggingContext() : base("RIGGINGContext")
        {
        }


        public DbSet<RiggingHeaderDS> RiggingHeaders;

        public DbSet<RiggingLinesDS> RiggingLines;


    }

}
