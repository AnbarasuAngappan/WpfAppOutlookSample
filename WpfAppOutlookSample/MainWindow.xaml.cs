using MacroView.VSTO.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using MacroView.VSTO.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WpfAppOutlookSample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Outlook.Application application = new Outlook.Application();

        public MainWindow()
        {
            InitializeComponent();
            SendEmailtoContacts();
        }
       

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {           

        }      
        
        private void SendEmailtoContacts()
        {
            

            string subjectEmail = "Meeting has been rescheduled.";
            string bodyEmail = "Meeting is one hour later.";

            Outlook.MAPIFolder sentContacts = (Outlook.MAPIFolder)
                application.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderContacts);


            foreach (Outlook.ContactItem contact in sentContacts.Items)
            {
                if (contact.Email1Address.Contains("example.com"))
                {
                    this.CreateEmailItem(subjectEmail, contact
                        .Email1Address, bodyEmail);
                }
            }

        }


        private void CreateEmailItem(string subjectEmail,string toEmail, string bodyEmail)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)
                application.CreateItem(Outlook.OlItemType.olMailItem);

            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }


        private void CreateInboxSubFolder(Outlook.Application OutlookApp)
        {
            Outlook.NameSpace nameSpace = OutlookApp.GetNamespace("MAPI");
            Outlook.MAPIFolder folderInbox = nameSpace.GetDefaultFolder(
                  Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Folders inboxFolders = folderInbox.Folders;
            Outlook.MAPIFolder subfolderInbox = null;
            try
            {
                subfolderInbox = inboxFolders.Add("InboxSubfolder",
                     Outlook.OlDefaultFolders.olFolderInbox);
            }
            catch (COMException exception)
            {
                if (exception.ErrorCode == -2147352567)
                    //  Cannot create the folder.
                   MessageBox.Show(exception.Message);
            }
            if (subfolderInbox != null) Marshal.ReleaseComObject(subfolderInbox);
            if (inboxFolders != null) Marshal.ReleaseComObject(inboxFolders);
            if (folderInbox != null) Marshal.ReleaseComObject(folderInbox);
            if (nameSpace != null) Marshal.ReleaseComObject(nameSpace);
        }

        private void AccessContacts(string findLastName)
        {


            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Outlook.Items items = (Outlook.Items)inBox.Items;


            //Outlook.Folder oFolder = (Outlook.Folder)olNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFoderDrafts);

            //Outlook.MAPIFolder folderContacts = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

            //Outlook.Items searchFolder = folderContacts.Items;
            //int counter = 0;

            //foreach (Outlook.ContactItem foundContact in searchFolder)
            //{
            //    if (foundContact.LastName.Contains(findLastName))
            //    {
            //        foundContact.Display(false);
            //        counter = counter + 1;
            //    }
            //}

            //MessageBox.Show("You have " + counter +
            //    " contacts with last names that contain "
            //    + findLastName + ".");
        }
    }
}
