using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OutlookAddIn_Report_Spam
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReportBtn_Click(object sender, RibbonControlEventArgs e)
        {
            MailHelper.CreateMail_SelectedMails_Attached();
        }

    }
}
