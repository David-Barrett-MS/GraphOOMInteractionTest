using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Net.Http;

namespace GraphOOMInteractionTest
{
    internal class GraphWatcher
    {
        // In a production application, we would use a push notification to get notified of new messages.
        // We don't do that here as we want control over the timing.
        private string appId;
        private string appSecret;
        private string tenantId;
        private string mailbox;
        private IConfidentialClientApplication graphApp;
        private AuthenticationResult lastAuthResult = null;
        private System.Timers.Timer messageCheckTimer = new System.Timers.Timer(10000);
        public string messageToDeleteSubject = "";
        private HttpClient httpClient = new HttpClient();

        internal GraphWatcher(string AppId, string AppSecret, string TenantId, string Mailbox)
        {
            appId = AppId;
            appSecret = AppSecret;
            tenantId = TenantId;
            mailbox = Mailbox;

            graphApp = ConfidentialClientApplicationBuilder.Create(appId)
                .WithClientSecret(appSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}")).Build();
            Task.Run(GetAccessToken).Wait();

            messageCheckTimer.Elapsed += MessageCheckTimer_Elapsed;
            messageCheckTimer.Start();
        }

        public void CheckForMessageToDelete()
        {
            Console.WriteLine($"GRAPH - Checking for message to delete: {messageToDeleteSubject}");

            string searchMessageUrl = "https://graph.microsoft.com/v1.0/users/" + mailbox + "/messages?$filter=subject eq '" + messageToDeleteSubject + "'"; // /mailfolder/inbox
            HttpResponseMessage response = httpClient.GetAsync(searchMessageUrl).Result;

            if (response.IsSuccessStatusCode)
            {
                string responseContent = response.Content.ReadAsStringAsync().Result;
                if (responseContent.Contains("value\":[]"))
                {
                    //Console.WriteLine($"GRAPH - Message not found: {messageToDeleteSubject}");
                }
                else
                {
                    Console.WriteLine($"GRAPH - Message found: {messageToDeleteSubject}");
                    // Delete the message
                    string messageId = responseContent.Split(new string[] { "\"id\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    string deleteMessageUrl = "https://graph.microsoft.com/v1.0/users/" + mailbox + "/messages/" + messageId;
                    response = httpClient.DeleteAsync(deleteMessageUrl).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"GRAPH - Message deleted: {messageToDeleteSubject}");
                    }
                    else
                    {
                        Console.WriteLine($"GRAPH - Error deleting message: {messageToDeleteSubject}");
                    }
                    messageToDeleteSubject = "";
                }
            }
            else
            {
                Console.WriteLine($"GRAPH - Error searching for message: {messageToDeleteSubject}");
            }

        }

        private void MessageCheckTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (String.IsNullOrEmpty(messageToDeleteSubject)) return;
            messageCheckTimer.Stop();
            CheckForMessageToDelete();
            messageCheckTimer.Start();
        }

        private async Task<string> GetAccessToken()
        {
            lastAuthResult = await graphApp.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            Console.WriteLine($"GRAPH - Got access token (expires {lastAuthResult.ExpiresOn})");
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + lastAuthResult.AccessToken);
            return lastAuthResult.AccessToken;
        }

        public bool SendMessage(string senderMailbox)
        {
            // Send a message to the specified mailbox
            // Invoke-RestMethod -Method Post -Uri $createUrl -Headers $authHeader -Body $sendMessageJson -ContentType "application/json"
            
            // Prepare the URL
            string sendMessageUrl = "https://graph.microsoft.com/v1.0/users/" + senderMailbox + "/sendMail";

            // Prepare the message JSON
            string subject = $"{DateTime.Now}-OOMDELETE";
            string sendMessageJson = SendMessageJSON(subject, mailbox);

            // Send the message
            HttpResponseMessage response = httpClient.PostAsync(sendMessageUrl, new StringContent(sendMessageJson, Encoding.UTF8, "application/json")).Result;
            if (response.IsSuccessStatusCode)
            {
                messageToDeleteSubject = subject;
                return true;
            }
            return false;
        }

        private string SendMessageJSON(string subject, string recipient)
        {
            StringBuilder sendMessageJSON = new StringBuilder("{\r\n  \"message\": {\r\n    \"subject\": \"");
            sendMessageJSON.Append(subject);
            sendMessageJSON.Append("\",\r\n    \"body\": {\r\n      \"contentType\": \"Text\",\r\n      \"content\": \"Test for sync.\"\r\n    },\r\n    \"toRecipients\": [\r\n      {\r\n        \"emailAddress\": {\r\n          \"address\": \"");
            sendMessageJSON.Append(recipient);
            sendMessageJSON.Append("\"\r\n        }\r\n      }\r\n    ]\r\n  },\r\n  \"saveToSentItems\": \"false\"\r\n}");
            return sendMessageJSON.ToString();
        }
    }
}
