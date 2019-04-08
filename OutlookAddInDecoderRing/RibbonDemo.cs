using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInDecoderRing
{
    public partial class RibbonDemo
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDecodeMessage_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the Application object
            Outlook.Application application = Globals.ThisAddIn.Application;

            // Get the active Inspector object and check if is type of MailItem
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                MessageBox.Show("Subject: " + mailItem.Subject);
            }
        }
    }
}
