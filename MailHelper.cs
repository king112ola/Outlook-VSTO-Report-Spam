using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_Report_Spam
{

    public class MailHelper
    {
        static public Outlook.Explorer currentExplorer = null;
        static private Outlook.Application currentApplication = Globals.ThisAddIn.Application;
        static public Dictionary<string, object> customizedRegistryValue = Globals.ThisAddIn.customizedRegistryValue;

        // Parent Function that invokes the helper / handlers
        public static void CreateMail_SelectedMails_Attached()
        {
            IEnumerable<MailItem> mailitems = MailHelper.GetSelectedEmails_Handler();

            if (MailItemsStatusCheck_Helper(mailitems) == false) return;

            Outlook.MailItem newMail = currentApplication.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

            Outlook.MailItem newMailWithAttachments = AddMailItemsAsAttachment_Handler(mailitems, newMail);

            // Test to display the email window, restricted to Development purpose
            // newMail.Display(false);

            Boolean sendMailConfirmation = SendMailConfirmation_Helper();

            if (sendMailConfirmation) newMail.Send();
        }

        /**
         * Handler for composing a new mail with mailitems as attachments
         * 
         * \mailitems IEnumerable<MailItem> instance contains iterative mail items.
         * \newMail Empty new mail item to be attached with mailitems.
         */
        public static MailItem AddMailItemsAsAttachment_Handler(IEnumerable<MailItem> mailitems, MailItem newMail)
        {
            foreach (MailItem mailitem in mailitems)
            {
                newMail.Attachments.Add(mailitem, Outlook.OlAttachmentType.olEmbeddeditem, 1, mailitem.Subject);
                newMail.To = (string)customizedRegistryValue["SOCEmail"];
                newMail.Subject = (string)customizedRegistryValue["EmailSubject"];
                newMail.HTMLBody = (string)customizedRegistryValue["EmailBody"];
            }
            return newMail;
        }

        // Simple handler to retrieve mail selected by user
        public static IEnumerable<MailItem> GetSelectedEmails_Handler()
        {

            foreach (Object mail_obj in currentApplication.ActiveExplorer().Selection)
            {
                if (mail_obj is MailItem)
                {
                    yield return (MailItem)mail_obj;

                }
            }


        }

        /**
         * Helper function for asserting if the MailItems has passed a conditional check
         * 
         * \return a simple Boolean which true indicates pass and false equals block
         */
        static public Boolean MailItemsStatusCheck_Helper(IEnumerable<MailItem> mailitems)
        {
            // Selected Mail is empty
            if (mailitems.Count() == 0) return false;

            return true;

        }

        // Confirmation dialog to determine whether the user agrees to send
        static public Boolean SendMailConfirmation_Helper()
        {
            DialogResult dialogResult = MessageBox.Show("Confirm Report Spam Email?", "Report Spam", MessageBoxButtons.OKCancel);

            if (dialogResult == DialogResult.OK) return true;

            return false;
        }

    }
}
