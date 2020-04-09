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

            var fileAttachments = await GetFileAttachments(graphClient, messages);
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
            var msgIdToAttachmentsDict = new Dictionary<Message, List<FileAttachment>>();
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
                    msgIdToAttachmentsDict[messages[i]] = attachmentsList;
                }
            }
            return msgIdToAttachmentsDict;
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
    }
}