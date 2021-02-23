using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Text;

namespace BadOutlook
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; }

        public string EmailSubject { get; set; }

        public string EmailBody { get; set; }

        public static List<OutlookEmails> ReadMailItems()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;

            try
            {

                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;

                foreach (Object item in mailItems)
                {
                    if (item is MailItem)
                    {
                        emailDetails = new OutlookEmails();
                        emailDetails.EmailFrom = ((MailItem) item).SenderEmailAddress;
                        emailDetails.EmailSubject = ((MailItem)item).Subject;
                        emailDetails.EmailBody = ((MailItem)item).Body;
                        listEmailDetails.Add(emailDetails);
                        ReleaseComObject(item);
                    }
                }

            }
            catch (System.Exception ex)
            {

                Console.WriteLine(ex.Message);

            }
            finally
            {

                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);

            }

            return listEmailDetails;

        }

        private static void ReleaseComObject(object obj)
        {

            if (obj != null)
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;

            }

        }

    }
}
