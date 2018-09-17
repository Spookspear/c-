using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;


namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {

        public void ClassTestCodeMotherChild(Excel.Application xls)
        {
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet WksMaster = Wkb.ActiveSheet;                    // get the sheet we are on  // point to active sheet

            // delare lists
            List<Mother> lstMothers = new List<Mother>();
            List<Child> lstChildren = new List<Child>();
            Mother clsMother = new Mother();
            Child clsChild = new Child();

            // pass to a seperate procedure to populate
            populateMotherChildObjects(lstMothers, lstChildren, clsMother, clsChild);
            populateWorksheetFromMotherChildObjects(WksMaster, lstMothers);


        }

        private void populateWorksheetFromMotherChildObjects(Worksheet wksMaster, List<Mother> lstMothers)
        {
            int iRow = 2;
            int iCol = 1;

            foreach (Mother m in lstMothers)
            {
                wksMaster.Cells[iRow, iCol].value = m.FullName();
                iCol++;
                wksMaster.Cells[iRow, iCol].value = m.Age;
                iCol++;

                foreach (Child c in m.lstChildren)
                {
                    wksMaster.Cells[iRow, iCol].value = c.FullName();
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = c.Age;


                    iRow++;
                    iCol = 3;
                }
                iRow++;
                iCol = 1;
            }

        }

        private void populateMotherChildObjects(List<Mother> lstMothers, List<Child> lstChildren, Mother clsMother, Child clsChild)
        {

            // instantiate objects from classes

            // 1st mother
            lstChildren = new List<Child>();

            // clsMother.Name = "Mandy";
            clsMother.FirstName = "Mandy";
            clsMother.LastName = "Bishop";
            clsMother.Age = 45;
            // out of intrest?
            clsMother.FullName();


            clsChild.FirstName = "Anthony";
            clsChild.LastName = "Bishop";
            clsChild.Age = 16;

            lstChildren.Add(clsChild);

            clsChild = new Child();
            clsChild.FirstName = "Katie";
            clsChild.LastName = "Bishop";
            clsChild.Age = 14;

            lstChildren.Add(clsChild);

            // assign children to the mother
            clsMother.lstChildren = lstChildren;
            lstMothers.Add(clsMother);

            // ---------------------------------------------------------------------------------------------------------
            // 2nd mother
            // ---------------------------------------------------------------------------------------------------------
            clsMother = new Mother();
            clsMother.FirstName = "Carol";
            clsMother.LastName = "Bishop";
            clsMother.Age = 71;
            // clsMother.FullName();

            lstChildren = new List<Child>();

            clsChild = new Child();
            clsChild.FirstName = "Grant";
            clsChild.LastName = "Bishop";
            clsChild.Age = 52;
            lstChildren.Add(clsChild);

            clsChild = new Child();
            clsChild.FirstName = "Jason";
            clsChild.LastName = "Bishop";
            clsChild.Age = 50;

            lstChildren.Add(clsChild);

            clsChild = new Child();
            clsChild.FirstName = "Jayne";
            clsChild.LastName = "Bishop";
            clsChild.Age = 47;

            lstChildren.Add(clsChild);

            clsMother.lstChildren = lstChildren;
            lstMothers.Add(clsMother);

            // end of 2nd

            // ---------------------------------------------------------------------------------------------------------
            // 3rd mother
            // ---------------------------------------------------------------------------------------------------------
            clsMother = new Mother();
            clsMother.FirstName = "Mary";
            clsMother.LastName = "Williamson";
            clsMother.Age = 91;
            // clsMother.FullName();

            lstChildren = new List<Child>();

            clsChild = new Child();
            clsChild.FirstName = "Frank";
            clsChild.LastName = "Williamson";
            clsChild.Age = 70;

            lstChildren.Add(clsChild);

            clsChild = new Child();
            clsChild.FirstName = "Shirley";
            clsChild.LastName = "Williamson";
            clsChild.Age = 65;

            lstChildren.Add(clsChild);

            clsMother.lstChildren = lstChildren;
            lstMothers.Add(clsMother);

        }

    }
}
