using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections;
using Microsoft.Win32;
using System.Diagnostics;

namespace OutlookAddIn_Report_Spam
{
    public partial class ThisAddIn
    {

        public MailHelper mailHelper = null;

        public Dictionary<string, object> customizedRegistryValue = new Dictionary<string, object>();

        const string RegisterSubkey = "Software\\Microsoft\\Office\\Outlook\\AddinsData\\ReportSpamMail";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            customizedRegistryValue = getCustomizedRegistryValue();
            if (!customizedRegistryValue.ContainsKey("SOCEmail")) initCustomizedRegistryValue();
        }

        private Dictionary<string, object> getCustomizedRegistryValue()
        {

            Dictionary<string, object> _customizedRegistryValue = new Dictionary<string, object>();

            RegistryKey registryKey = getHKCURegistryKey(RegisterSubkey);

            if (registryKey != null)
            {
                string[] valueNames = registryKey.GetValueNames();
                foreach (string currentKey in valueNames)
                {
                    object tmpValue = registryKey.GetValue(currentKey);
                    _customizedRegistryValue.Add(currentKey, tmpValue);
                }
                registryKey.Close();

            }

            return _customizedRegistryValue;
        }

        static private RegistryKey getHKCURegistryKey(string keyName) => Registry.CurrentUser.OpenSubKey(RegisterSubkey, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);

        static private void initCustomizedRegistryValue()
        {
            Dictionary<string, object> keyValueNames = new Dictionary<string, object>()
            {
                ["EmailBody"] = "This is an automated email reporting a spam email received by a user.<br/>These unsolicited emails may contains suspicious content and displays clear signs of deceptive intent.<br/>Please review the attached Spam Emails for further investigation.",
                ["EmailSubject"] = "[Report Spam] Potential Spam Email Report From User",
                ["SOCEmail"] = "david.li@amidas.life",      
            };

            RegistryKey registryKey = getHKCURegistryKey(RegisterSubkey) ?? Registry.CurrentUser.CreateSubKey(RegisterSubkey);

            foreach (var keyValueName in keyValueNames)
                {
                registryKey.SetValue((string)keyValueName.Key, (string)keyValueName.Value);
                }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
