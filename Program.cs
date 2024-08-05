using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Text;

namespace GraphOOMInteractionTest
{
    internal class Program
    {
        static internal OutlookWatcher outlookWatcher;
        static internal GraphWatcher graphWatcher;
        static bool eventBasedGraphDelete = false;

        static void Main(string[] args)
        {
            ShowHelp();

            outlookWatcher = new OutlookWatcher();
            outlookWatcher.ItemDeleted += OutlookWatcher_ItemDeleted;
            graphWatcher = new GraphWatcher(Properties.Settings.Default.AppId, Properties.Settings.Default.SecretKey, Properties.Settings.Default.TenantId, Properties.Settings.Default.Mailbox);

            ConsoleKeyInfo key = Console.ReadKey();
            while (key.Key != ConsoleKey.X)
            {
                if (key.Key == ConsoleKey.S)
                {
                    Console.WriteLine("Sending message...");
                    if (graphWatcher.SendMessage(Properties.Settings.Default.SenderMailbox))
                        Console.WriteLine("Message sent");
                    else
                        Console.WriteLine("Message failed to send");
                }
                else if (key.Key == ConsoleKey.E)
                {
                    eventBasedGraphDelete = !eventBasedGraphDelete;
                    Console.WriteLine($"Event-based Graph delete is now: {eventBasedGraphDelete}");
                }
                else if (key.Key == ConsoleKey.H)
                {
                    ShowHelp();
                }
                key = Console.ReadKey();
            }
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Available keys:");
            Console.WriteLine("H - show this help");
            Console.WriteLine($"S - send a message (using Graph, from {Properties.Settings.Default.SenderMailbox}");
            Console.WriteLine("E - toggle event-based Graph delete (Graph delete is triggered directly after Outlook delete call and is likely to demonstrate timing issues)");
            Console.WriteLine("X - exit");
            Console.WriteLine();
        }

        private static void OutlookWatcher_ItemDeleted(object sender, EventArgs e)
        {
            if (eventBasedGraphDelete)
                graphWatcher.CheckForMessageToDelete();
        }
    }
}
