using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInDecoderRing
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors _inspectors;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += Inspectors_NewInspector;

            Outlook.Application application = Application;
            application.ItemSend += ItemSend_BeforeSend;

            // Get the active Inspector object
            Outlook.Inspector activeInspector = application.ActiveInspector();
            if (activeInspector != null)
            {
                // Get the title of the active item when the Outlook start.
                MessageBox.Show("Active inspector: " + activeInspector.Caption);
            }

            // Get the active Explorer object
            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                // Get the title of the active folder when the Outlook start.
                //MessageBox.Show("Active explorer: " + activeExplorer.Caption);
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
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }


        void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }
            }
        }

        void ItemSend_BeforeSend(object item, ref bool cancel)
        {
            // encode the message here
            Outlook.MailItem mailItem = (Outlook.MailItem)item;
            if (mailItem != null)
            {
                mailItem.Body += "Modified by GettingStartedOutlookAddIn";
            }
            cancel = false;
        }
        #endregion
    }
}
