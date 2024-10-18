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

using System;

namespace GraphOOMInteractionTest
{
    internal class Program
    {
        static internal OutlookWatcher _outlookWatcher;
        static internal GraphWatcher _graphWatcher;
        static bool _eventBasedGraphDelete = false;

        static void Main(string[] args)
        {
            ShowHelp();

            _outlookWatcher = new OutlookWatcher();
            _outlookWatcher.ItemDeleted += OutlookWatcher_ItemDeleted;
            _graphWatcher = new GraphWatcher(Properties.Settings.Default.AppId, Properties.Settings.Default.SecretKey, Properties.Settings.Default.TenantId, Properties.Settings.Default.Mailbox);

            ConsoleKeyInfo key = Console.ReadKey(true);
            while (key.Key != ConsoleKey.X)
            {
                switch (key.Key)
                {
                    case ConsoleKey.C:
                        Console.WriteLine($"Current time interval for Graph item search: {_graphWatcher.CheckInterval}");
                        Console.WriteLine("Enter new time interval for Graph item search (in milliseconds):");
                        string newInterval = Console.ReadLine();
                        if (double.TryParse(newInterval, out double interval))
                        {
                            _graphWatcher.CheckInterval = interval;
                        }
                        else
                            Console.WriteLine("Invalid input");
                        Console.WriteLine($"Current time interval for Graph item search: {_graphWatcher.CheckInterval}");
                        break;

                    case ConsoleKey.S:
                        Console.WriteLine("Sending message...");
                        if (_graphWatcher.SendMessage(Properties.Settings.Default.SenderMailbox))
                            Console.WriteLine("Message sent");
                        else
                            Console.WriteLine("Message failed to send");
                        break;

                    case ConsoleKey.E:
                        _eventBasedGraphDelete = !_eventBasedGraphDelete;
                        ShowOOMDeleteTriggerStatus();
                        break;

                    default:
                        ShowHelp();
                        break;
                }

                key = Console.ReadKey(true);
            }
        }

        private static void ShowOOMDeleteTriggerStatus()
        {
            Console.Write($"Graph delete triggered at same time as OOM delete: ");
            if (_eventBasedGraphDelete)
                Console.WriteLine("Enabled");
            else
                Console.WriteLine("Disabled");
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Available keys:");
            Console.WriteLine("H - show this help");
            Console.WriteLine($"S - send a message (using Graph, from {Properties.Settings.Default.SenderMailbox}");
            Console.WriteLine("C - change the time interval for Graph item search");
            Console.WriteLine("E - toggle OOM event-based Graph delete (Graph delete is sent directly after Outlook delete call)");
            Console.WriteLine("X - exit");
            Console.WriteLine();
            if (_graphWatcher != null)
                Console.WriteLine($"Current time interval for Graph item search: {_graphWatcher.CheckInterval}");
            ShowOOMDeleteTriggerStatus();
            Console.WriteLine();
        }

        private static void OutlookWatcher_ItemDeleted(object sender, EventArgs e)
        {
            if (_eventBasedGraphDelete)
                _graphWatcher.CheckForMessageToDelete();
        }
    }
}
