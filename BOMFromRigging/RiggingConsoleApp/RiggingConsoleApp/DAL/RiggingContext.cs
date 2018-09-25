using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiggingConsoleApp.DAL
{
    class RiggingContext : DbContext
    {
        public RiggingContext() : base("RiggingContext") { }

        public DbSet<RiggingConsoleApp.DAL.Models.RiggingHeaderDS> RiggingHeaders { get; set; }

         
    }
}
