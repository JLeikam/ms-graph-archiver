using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ms_graph_app.Models;
using Newtonsoft.Json;
using System.Net;
using System.Threading;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.PixelFormats;
using Image = SixLabors.ImageSharp.Image;
using KeyValuePair = System.Collections.Generic.KeyValuePair;
using System.Web;
using Newtonsoft.Json.Linq;

namespace ms_graph_app.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MessagesController : ControllerBase
    {
        private readonly GraphConfig config;
        private static Dictionary<string, Subscription> Subscriptions = new Dictionary<string, Subscription>();
        private static Timer subscriptionTimer = null;

        public MessagesController(GraphConfig config)
        {
            this.config = config;
        }

        [HttpGet]
        public async Task<ActionResult<string>> Get()
        {
            var graphServiceClient = GetGraphClient();
            var sub = new Microsoft.Graph.Subscription();
            sub.ChangeType = "created";
            sub.NotificationUrl = $"{config.Ngrok}/api/messages";
            sub.Resource = $"/users/jleikam@integrativemeaning.com/mailFolders/{config.ArchiverId}/messages?$filter=isRead eq false";
            sub.ExpirationDateTime = DateTime.UtcNow.AddMinutes(5);
            sub.ClientState = Guid.NewGuid().ToString();

            var newSubscription = await graphServiceClient
                .Subscriptions
                .Request()
                .AddAsync(sub);

            Subscriptions[newSubscription.Id] = newSubscription;

            if(subscriptionTimer == null)
            {
                subscriptionTimer = new Timer(CheckSubscriptions, null, 5000, 15000);
            }

            return $"Subscribed. Id: {newSubscription.Id}, Expiration: {newSubscription.ExpirationDateTime}";
        }

        public async Task<ActionResult<string>> Post([FromQuery]string validationToken = null)
        {
            // handle validation
            if (!string.IsNullOrEmpty(validationToken))
            {
                Console.WriteLine($"Received Token: '{validationToken}'");
                return Ok(validationToken);
            }

            // handle notifications
            using (StreamReader reader = new StreamReader(Request.Body))
            {
                string content = await reader.ReadToEndAsync();

                Console.WriteLine(content);

                var notifications = JsonConvert.DeserializeObject<Notifications>(content);

                foreach (var notification in notifications.Items)
                {
                    Console.WriteLine($"Received notification: '{notification.Resource}', {notification.ResourceData?.Id}");
                }
            }

            // query for updates
            await CheckForUpdates();

            return Ok();
        }

        private GraphServiceClient GetGraphClient()
        {
            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                // get an access token for Graph
                var accessToken = GetAccessToken().Result;

                requestMessage
                    .Headers
                    .Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                return Task.FromResult(0);
            }));

            return graphClient;
        }

        private async Task<string> GetAccessToken()
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(config.AppId)
              .WithClientSecret(config.AppSecret)
              .WithAuthority($"https://login.microsoftonline.com/{config.TenantId}")
              .WithRedirectUri("https://daemon")
              .Build();

            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            return result.AccessToken;
        }

        private void CheckSubscriptions(Object stateInfo)
        {
            AutoResetEvent autoEvent = (AutoResetEvent)stateInfo;

            Console.WriteLine($"Checking subscriptions {DateTime.Now.ToString("h:mm:ss.fff")}");
            Console.WriteLine($"Current subscription count {Subscriptions.Count()}");

            foreach (var subscription in Subscriptions)
            {
                // if the subscription expires in the next 2 min, renew it
                if (subscription.Value.ExpirationDateTime < DateTime.UtcNow.AddMinutes(2))
                {
                    RenewSubscription(subscription.Value);
                }
            }
        }

        private async void RenewSubscription(Subscription subscription)
        {
            Console.WriteLine($"Current subscription: {subscription.Id}, Expiration: {subscription.ExpirationDateTime}");

            var graphServiceClient = GetGraphClient();

            var newSubscription = new Subscription
            {
                ExpirationDateTime = DateTime.UtcNow.AddMinutes(5)
            };

            await graphServiceClient
              .Subscriptions[subscription.Id]
              .Request()
              .UpdateAsync(newSubscription);

            subscription.ExpirationDateTime = newSubscription.ExpirationDateTime;
            Console.WriteLine($"Renewed subscription: {subscription.Id}, New Expiration: {subscription.ExpirationDateTime}");
        }

        
        private async Task CheckForUpdates()
        {
            var graphClient = GetGraphClient();

            IMailFolderMessagesCollectionPage messages = await GetUnreadMessages(graphClient);

            var msgToAttachmentsDict = await GetFileAttachments(graphClient, messages);


            foreach(KeyValuePair<Message, List<FileAttachment>> kvp in msgToAttachmentsDict)
            {
                foreach(FileAttachment attachment in kvp.Value)
                {
                    if (attachment.ContentType.Contains("image"))
                    {
                        string txt = await OCR(attachment.ContentBytes);
                    }
                    else if (attachment.ContentType.Contains("csv"))
                    {
                        //csv parser
                    }
                }
            }
            //TODO: add marked read when attachments have been retrieved

            //if (base64str != null)
            //{
            //    using (var client = new HttpClient())
            //    {
            //        string[] base64StrArr = new string[] { base64str };
            //        client.BaseAddress = new Uri("http://localhost:5000");
            //        var response = client.PostAsJsonAsync("/api/vision", base64StrArr).Result;
            //    }
            //}

            OutputMessages(messages);

        }

        private async Task<Dictionary<Message, List<FileAttachment>>> GetFileAttachments(GraphServiceClient graphClient, IMailFolderMessagesCollectionPage messages)
        {
            var msgToAttachmentsDict = new Dictionary<Message, List<FileAttachment>>();
            for(int i = 0; i<messages.Count; i++)
            {
                if (messages[i].HasAttachments == true)
                {
                    var attachmentsList = new List<FileAttachment>();
                    IMessageAttachmentsCollectionPage attachmentsPage = await graphClient.Users["jleikam@integrativemeaning.com"]
                                        .MailFolders
                                        [config.ArchiverId]
                                        .Messages[messages[i].Id]
                                        .Attachments
                                        .Request()
                                        .GetAsync();

                    for (int j = 0; j < attachmentsPage.Count; j++) {
                        if (attachmentsPage[j].ODataType == "#microsoft.graph.fileAttachment")
                        {
                            var fileAttachment = attachmentsPage[j] as FileAttachment;
                            attachmentsList.Add(fileAttachment);

                        }
                    }
                    msgToAttachmentsDict[messages[i]] = attachmentsList;
                }
            }
            return msgToAttachmentsDict;
        }

        private void OutputMessages(IMailFolderMessagesCollectionPage messages)
        {
            foreach (var message in messages)
            {
                var output = $"Message: {message.Id} {message.Subject} {message.Body}";
                Console.WriteLine(output);
            }
        }

        private async Task<IMailFolderMessagesCollectionPage> GetUnreadMessages(GraphServiceClient graphClient)
        {
            IMailFolderMessagesCollectionPage page;
                page = await graphClient.Users["jleikam@integrativemeaning.com"]
                                        .MailFolders
                                        [config.ArchiverId]
                                        .Messages
                                        .Request()
                                        .Filter("isRead eq false")
                                        .GetAsync();
            return page;
        }
        private async Task<string> OCR(byte[] contentBytes)
        { 
            try
            {
                HttpClient client = new HttpClient();

                // Request headers.
                client.DefaultRequestHeaders.Add(
                    "Ocp-Apim-Subscription-Key", config.SubscriptionKey);

                var builder = new UriBuilder(config.Endpoint);
                builder.Port = -1;
                var query = HttpUtility.ParseQueryString(builder.Query);
                query["language"] = "en";
                builder.Query = query.ToString();
                string url = builder.ToString();

                HttpResponseMessage response;

                // Two REST API methods are required to extract text.
                // One method to submit the image for processing, the other method
                // to retrieve the text found in the image.

                // operationLocation stores the URI of the second REST API method,
                // returned by the first REST API method.
                string operationLocation;


                // Adds the byte array as an octet stream to the request body.
                using (ByteArrayContent content = new ByteArrayContent(contentBytes))
                {
                    // This example uses the "application/octet-stream" content type.
                    // The other content types you can use are "application/json"
                    // and "multipart/form-data".
                    content.Headers.ContentType =
                        new MediaTypeHeaderValue("application/octet-stream");

                    // The first REST API method, Batch Read, starts
                    // the async process to analyze the written text in the image.
                    response = await client.PostAsync(url, content);
                }

                // The response header for the Batch Read method contains the URI
                // of the second method, Read Operation Result, which
                // returns the results of the process in the response body.
                // The Batch Read operation does not return anything in the response body.
                if (response.IsSuccessStatusCode)
                    operationLocation =
                        response.Headers.GetValues("Operation-Location").FirstOrDefault();
                else
                {
                    // Display the JSON error data.
                    string errorString = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("\n\nResponse:\n{0}\n",
                        JToken.Parse(errorString).ToString());
                    return "error";
                }

                // If the first REST API method completes successfully, the second 
                // REST API method retrieves the text written in the image.
                //
                // Note: The response may not be immediately available. Text
                // recognition is an asynchronous operation that can take a variable
                // amount of time depending on the length of the text.
                // You may need to wait or retry this operation.
                //
                // This example checks once per second for ten seconds.
                string contentString;
                int i = 0;
                do
                {
                    System.Threading.Thread.Sleep(1000);
                    response = await client.GetAsync(operationLocation);
                    contentString = await response.Content.ReadAsStringAsync();
                    ++i;
                }
                while (i < 60 && contentString.IndexOf("\"status\":\"succeeded\"") == -1);

                if (i == 60 && contentString.IndexOf("\"status\":\"succeeded\"") == -1)
                {
                    Console.WriteLine("\nTimeout error.\n");
                    return "";
                }



                var text = "";
                //Parse the data
                JObject myJson = JsonConvert.DeserializeObject<JObject>(contentString);

                foreach (JObject readResult in (JArray)myJson["analyzeResult"]["readResults"])
                {
                    foreach (JObject line in (JArray)readResult["lines"])
                    {
                        Console.WriteLine("Line: {0}", line.GetValue("text"));
                        text += " " + line.GetValue("text");
                    }

                }
                Console.WriteLine("Full text: {0}", text);
                return text;
                // Display the JSON response.
                //Console.WriteLine("\nResponse:\n\n{0}\n",
                //    JToken.Parse(contentString).ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("\n" + e.Message);
            }
            return "";
        }
    }
}