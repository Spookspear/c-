#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolbarOfFunctions
{
    public class Mother
    {

        // public string FullName { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public int Age { get; set; }

        public string FullName()
        {
            return FirstName.ToString() + ' ' + LastName.ToString();
        }


        public List<Child> lstChildren { get; set; }
                
    }

    public class Child
    {
    
        public string FirstName { get; set; }

        public string LastName { get; set; }

        public int Age { get; set; }

        public string FullName()
        {
            return FirstName.ToString() + ' ' + LastName.ToString();
        }


    }


}
