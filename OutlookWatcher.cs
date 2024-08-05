using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GraphOOMInteractionTest
{

    internal class OutlookWatcher
    {
        Outlook.Application outlookApp = new Outlook.Application();
        Outlook.Folder inboxFolder;

        internal delegate void ItemDeletedEventHandler(object sender, EventArgs e);
        internal event ItemDeletedEventHandler ItemDeleted;

        internal OutlookWatcher()
        {
            inboxFolder = outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            inboxFolder.Items.ItemAdd += Items_ItemAdd;
        }

        private void Items_ItemAdd(object Item)
        {
            if (!(Item is MailItem)) return;
            string subject = ((MailItem)Item).Subject;
            Console.WriteLine($"OUTLOOK - Item received: {subject}");

            if (subject.EndsWith("-OOMDELETE"))
            {
                // Delete the item (this is a move to Deleted Items)
                ((MailItem)Item).Delete();
                Console.WriteLine($"OUTLOOK - Item deleted: {subject}");
                OnItemDeleted(EventArgs.Empty);
            }
        }

        protected virtual void OnItemDeleted(EventArgs e)
        {
            ItemDeleted?.Invoke(this, e);
        }
    }
}
