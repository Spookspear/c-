#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using DaveChambers.FolderBrowserDialogEx;

using System.ComponentModel;
using System.Data;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Tools.Ribbon;

using ToolbarOfFunctions_CommonClasses;
using ToolbarOfFunctions_MyConstants;
using System.Runtime.InteropServices;

// using System.Data.SqlTypes;

using System.DirectoryServices.AccountManagement;

using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;

namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {

        public void loadADGroupIntoSpreadsheetActiveCell(Excel.Application xls)
        {
            CommonExcelClasses.MsgBox("loadADGroupIntoSpreadsheetActiveCell - is what this will do");


        }

        public void readGroupMembership(Excel.Application xls)
        {
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolTestCode = myData.TestCode;

            #endregion

            #region [Declare and instantiate variables for worksheet/book]
            // get worksheet name
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet Wks;   // get current sheet

            Wks = Wkb.ActiveSheet;

            string strGroupName = Wks.Name;

            #endregion

            #region [Ask to display a Message?]
            DialogResult dlgResult = DialogResult.Yes;
            string strMessage;

            if (boolDisplayInitialMessage)
            {
                strMessage = "Get active Directory Details for: " + strGroupName;

                if (booltimeTaken)
                    strMessage = strMessage + LF + " and display the time taken";

                dlgResult = MessageBox.Show(strMessage + "?", "Active Directory", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            #endregion


            #region [Start of work]
            if (dlgResult == DialogResult.Yes)
            {

                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                DateTime dteStart = DateTime.Now;

                zapWorksheet(Wkb);
                readGroupUserMembership(Wks, strGroupName);
                
                writeHeaders(Wks, "ADUsers", false);
                CommonExcelClasses.sortSheet(Wks,2);

                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);


                #region [Display Complete Message]
                if (boolDisplayCompleteMessage)
                {
                    strMessage = "";
                    strMessage = strMessage + "Complete ...";

                    if (booltimeTaken)
                    {
                        DateTime dteEnd = DateTime.Now;
                        int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                        strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                    }

                    CommonExcelClasses.MsgBox(strMessage);          // localisation?
                }
                #endregion





            }
            #endregion


        }


        public void readGroupUserMembership(Excel.Worksheet Wks, string strGroupName)
        {
            myData = myData.LoadMyData();               // read data from settings file

            bool boolTestCode = myData.TestCode;

            // can I used the same code for both?

            // Connection information
            // var connectionString = "LDAP://domain.com/DC=domain,DC=com";
            string connectionString = "LDAP://subsea7.net/DC=subsea7,DC=net";

            // Split the LDAP Uri
            var uri = new Uri(connectionString);
            var host = uri.Host;
            var container = uri.Segments.Count() >= 1 ? uri.Segments[1] : "";

            int intRow = 2;
            int intCol = 1;

            var princContext = new PrincipalContext(ContextType.Domain, host, container);

            if (boolTestCode)
            {
                // Create context to connect to AD

                // Get group
                GroupPrincipal qbeGroup = new GroupPrincipal(princContext, strGroupName);
                PrincipalSearcher srch = new PrincipalSearcher(qbeGroup);


                // find all matches
                foreach (var found in srch.FindAll())
                {
                    if (found is GroupPrincipal foundGroup)
                    {
                        // iterate over members
                        foreach (Principal user in foundGroup.GetMembers())
                        {
                            intCol = 1;

                            Wks.Cells[intRow, intCol].Value = user.SamAccountName;
                            intCol++;

                            Wks.Cells[intRow, intCol].Value = user.DisplayName;
                            intCol++;

                            Wks.Cells[intRow, intCol].Value = user.Description;
                            intCol++;


                            intRow++;

                        }

                    }

                }
            }
            else
            {

                // PrincipalContext princContext = new PrincipalContext(ContextType.Domain);
                GroupPrincipal group = GroupPrincipal.FindByIdentity(princContext, strGroupName);
                if (group != null)
                {
                    // iterate over members
                    foreach (Principal p in group.GetMembers())
                    {
                        Console.WriteLine("{0}: {1}", p.StructuralObjectClass, p.DisplayName);

                        // do whatever you need to do to those members
                        if (p is UserPrincipal User)
                        {

                            intCol = 1;

                            Wks.Cells[intRow, intCol].Value = User.SamAccountName;
                            intCol++;

                            Wks.Cells[intRow, intCol].Value = User.DisplayName;
                            intCol++;

                            Wks.Cells[intRow, intCol].Value = User.Description;
                            intCol++;

                            Wks.Cells[intRow, intCol].Value = User.IsAccountLockedOut();
                            intCol++;

                            intRow++;

                        }
                    }


                }

            }


        }

        public void readUsersGroupMembership(Excel.Application xls)
        {

            // vb code
            /*  Application.ScreenUpdating = False
                Call ZapWorkSheet(ActiveSheet, 1)           - DONE 
                Call readGroupMembership(ActiveSheet)       - DONE 
                Call writeHeaders(ActiveSheet, "USERS")     - DONE 
                Call sortSheet(ActiveSheet)
             */

            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolTestCode = myData.TestCode;

            #endregion

            #region [Declare and instantiate variables for worksheet/book]
            // get worksheet name
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet Wks;   // get current sheet

            Wks = Wkb.ActiveSheet;

            string strUserName = Wks.Name;

            #endregion

            #region [Ask to display a Message?]
            DialogResult dlgResult = DialogResult.Yes;
            string strMessage;

            if (boolDisplayInitialMessage)
            {
                strMessage = "Get active Directory Details for: " + strUserName;

                if (booltimeTaken)
                    strMessage = strMessage + LF + " and display the time taken";

                dlgResult = MessageBox.Show(strMessage + "?", "Active Directory", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            #endregion


            #region [Start of work]
            if (dlgResult == DialogResult.Yes)
            {

                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                DateTime dteStart = DateTime.Now;

                zapWorksheet(Wkb);
                readUsersGroupMembership(Wks, strUserName);
                writeHeaders(Wks, "ADGroups", false);
                CommonExcelClasses.sortSheet(Wks,1);

                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);


                #region [Display Complete Message]
                if (boolDisplayCompleteMessage)
                {
                    strMessage = "";
                    strMessage = strMessage + "Complete ...";

                    if (booltimeTaken)
                    {
                        DateTime dteEnd = DateTime.Now;
                        int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                        strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                    }

                    CommonExcelClasses.MsgBox(strMessage);          // localisation?
                }
                #endregion





            }
            #endregion



        }

        public void readUsersGroupMembership(Excel.Worksheet Wks, string strUserName)
        {

            myData = myData.LoadMyData();               // read data from settings file

            bool boolTestCode = myData.TestCode;

            // Connection information
            // var connectionString = "LDAP://domain.com/DC=domain,DC=com";
            string connectionString = "LDAP://subsea7.net/DC=subsea7,DC=net";

            // Split the LDAP Uri
            var uri = new Uri(connectionString);
            var host = uri.Host;
            var container = uri.Segments.Count() >= 1 ? uri.Segments[1] : "";

            // Create context to connect to AD
            var princContext = new PrincipalContext(ContextType.Domain, host, container);

            // Get User
            UserPrincipal user = UserPrincipal.FindByIdentity(princContext, IdentityType.SamAccountName, strUserName);

            int intRow = 2;
            int intCol = 1;
            // Browse user's groups
            foreach (GroupPrincipal group in user.GetGroups())
            {
                intCol = 1;

                Wks.Cells[intRow, intCol].Value = group.Name;
                intCol++;

                Wks.Cells[intRow, intCol].Value = group.Description;
                intCol++;

                Wks.Cells[intRow, intCol].Value = group.IsSecurityGroup;
                intCol++;

                Wks.Cells[intRow, intCol].Value = group.GroupScope.ToString();
                intCol++;

                intRow++;

                Console.Out.WriteLine(group.Name);

            }
        }

        public void readGroupMembership_001(Excel.Worksheet Wks, string strUserName)
        {
            Console.WriteLine("ListAllGroupsViaTokenGroups:");

            List<string> result = new List<string>();

            try
            {
                // result = ListAllGroupsViaTokenGroups("GBI01", "subsea7.net");

                int intRow = 2;

                foreach (var group in result)
                {
                    Console.WriteLine(group);
                    // WRITE OUT TO SHEET
                    Wks.Cells[intRow, 1].Value = group;

                    intRow++;


                }
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }

            Console.Read();


        }

        private static List<string> ListAllGroupsViaTokenGroups(string username, string domainName)
        {
            List<string> result = new List<string>();

            int i = 0;

            try
            {
                using (PrincipalContext domainContext = new PrincipalContext(ContextType.Domain, domainName))
                using (var searcher = new DirectorySearcher(new DirectoryEntry("LDAP://" + domainContext.Name)))
                {
                    searcher.Filter = String.Format("(&(objectClass=user)(sAMAccountName={0}))", username);
                    SearchResult sr = searcher.FindOne();

                    DirectoryEntry user = sr.GetDirectoryEntry();

                    // access to other user properties, via user.Properties["..."]

                    user.RefreshCache(new string[] { "tokenGroups" });

                    for (i = 0; i < user.Properties["tokenGroups"].Count; i++)
                    {
                        SecurityIdentifier sid = new SecurityIdentifier((byte[])user.Properties["tokenGroups"][i], 0);
                        NTAccount nt = (NTAccount)sid.Translate(typeof(NTAccount));

                        result.Add(nt.Translate(typeof(NTAccount)).ToString() + " (" + sid.ToString() + ")");
                    }
                }

                return result;
            }
            catch (Exception excpt)
            {
                CommonExcelClasses.MsgBox("There was a problem: " + excpt.Message + " line: " + i.ToString(), "Error");
                Console.WriteLine(excpt.Message);

                throw;
            }
        }

    }
}

