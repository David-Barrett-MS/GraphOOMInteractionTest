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
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Net.Http;

namespace GraphOOMInteractionTest
{
    internal class GraphWatcher
    {
        // In a production application, we would use a push/change notification to get notified of new messages.
        // We don't do that here as we want control over the timing.
        private readonly string _appId;
        private readonly string _appSecret;
        private readonly string _tenantId;
        private readonly string _mailbox;
        private IConfidentialClientApplication _graphApp;
        private AuthenticationResult _lastAuthResult = null;
        private readonly System.Timers.Timer _messageCheckTimer;
        private double _graphCheckInterval = 10000;
        private readonly HttpClient _httpClient = new HttpClient();

        public string MessageToDeleteSubject = "";
        public bool SingleDeleteOnly = true;

        internal GraphWatcher(string AppId, string AppSecret, string TenantId, string Mailbox)
        {
            _appId = AppId;
            _appSecret = AppSecret;
            _tenantId = TenantId;
            _mailbox = Mailbox;

            _graphApp = ConfidentialClientApplicationBuilder.Create(_appId)
                .WithClientSecret(_appSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{_tenantId}")).Build();
            Task.Run(GetAccessToken).Wait();

            _messageCheckTimer = new System.Timers.Timer(_graphCheckInterval);
            _messageCheckTimer.Elapsed += MessageCheckTimer_Elapsed;
            _messageCheckTimer.Start();
        }

        public double CheckInterval
        {
            get { return _graphCheckInterval; }
            set { SetCheckInterval(value); }
        }

        public void SetCheckInterval(double interval)
        {
            _graphCheckInterval = interval;
            _messageCheckTimer.Stop();
            _messageCheckTimer.Interval = interval;
            _messageCheckTimer.Start();
        }

        public void CheckForMessageToDelete()
        {
            Console.WriteLine($"GRAPH - Checking for message to delete: {MessageToDeleteSubject}");

            string searchMessageUrl = "https://graph.microsoft.com/v1.0/users/" + _mailbox + "/messages?$filter=subject eq '" + MessageToDeleteSubject + "'"; // /mailfolder/inbox
            HttpResponseMessage response = _httpClient.GetAsync(searchMessageUrl).Result;

            if (response.IsSuccessStatusCode)
            {
                string responseContent = response.Content.ReadAsStringAsync().Result;
                if (responseContent.Contains("value\":[]"))
                {
                    //Console.WriteLine($"GRAPH - Message not found: {messageToDeleteSubject}");
                }
                else
                {
                    Console.WriteLine($"GRAPH - Message found: {MessageToDeleteSubject}");
                    // Delete the message
                    string messageId = responseContent.Split(new string[] { "\"id\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    string deleteMessageUrl = "https://graph.microsoft.com/v1.0/users/" + _mailbox + "/messages/" + messageId;
                    response = _httpClient.DeleteAsync(deleteMessageUrl).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"GRAPH - Message deleted: {MessageToDeleteSubject}");
                    }
                    else
                    {
                        Console.WriteLine($"GRAPH - Error deleting message: {MessageToDeleteSubject}");
                    }
                    if (SingleDeleteOnly)
                        MessageToDeleteSubject = "";
                }
            }
            else
            {
                Console.WriteLine($"GRAPH - Error searching for message: {MessageToDeleteSubject}");
            }

        }

        private void MessageCheckTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (String.IsNullOrEmpty(MessageToDeleteSubject)) return;
            _messageCheckTimer.Stop();
            CheckForMessageToDelete();
            _messageCheckTimer.Start();
        }

        private async Task<string> GetAccessToken()
        {
            _lastAuthResult = await _graphApp.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            Console.WriteLine($"GRAPH - Got access token (expires {_lastAuthResult.ExpiresOn})");
            _httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + _lastAuthResult.AccessToken);
            return _lastAuthResult.AccessToken;
        }

        public bool SendMessage(string senderMailbox)
        {
            // Send a message to the specified mailbox
            // Invoke-RestMethod -Method Post -Uri $createUrl -Headers $authHeader -Body $sendMessageJson -ContentType "application/json"
            
            // Prepare the URL
            string sendMessageUrl = "https://graph.microsoft.com/v1.0/users/" + senderMailbox + "/sendMail";

            // Prepare the message JSON
            string subject = $"{DateTime.Now}-OOMDELETE";
            string sendMessageJson = SendMessageJSON(subject, _mailbox);

            // Send the message
            HttpResponseMessage response = _httpClient.PostAsync(sendMessageUrl, new StringContent(sendMessageJson, Encoding.UTF8, "application/json")).Result;
            if (response.IsSuccessStatusCode)
            {
                MessageToDeleteSubject = subject;
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
