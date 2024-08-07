/*
 * By David Barrett, Microsoft Ltd. 2024. Use at your own risk.  No warranties are given.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

using Microsoft.Office.Interop.Outlook;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GraphOOMInteractionTest
{

    internal class OutlookWatcher
    {
        Outlook.Application _outlookApp = new Outlook.Application();
        Outlook.Folder _inboxFolder;

        internal delegate void ItemDeletedEventHandler(object sender, EventArgs e);
        internal event ItemDeletedEventHandler ItemDeleted;

        internal OutlookWatcher()
        {
            _inboxFolder = _outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            _inboxFolder.Items.ItemAdd += Items_ItemAdd;
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
