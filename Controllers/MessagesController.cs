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
        private readonly MyConfig config;
        private static Dictionary<string, Subscription> Subscriptions = new Dictionary<string, Subscription>();
        private static Timer subscriptionTimer = null;

        public MessagesController(MyConfig config)
        {
            this.config = config;
        }

        [HttpGet]
        public async Task<ActionResult<string>> Get()
        {
            var graphServiceClient = GetGraphClient();
            var messagesEndpoint = graphServiceClient.Users["jleikam@integrativemeaning.com"].MailFolders.Inbox.Messages.RequestUrl;
            var sub = new Microsoft.Graph.Subscription();
            sub.ChangeType = "created";
            sub.NotificationUrl = $"{config.Ngrok}/api/messages";
            sub.Resource = "/users/jleikam@integrativemeaning.com/mailFolders/Inbox/messages";
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

            // use deltaquery to query for all updates
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

        private static object DeltaLink = null;

        private static IMessageDeltaCollectionPage lastPage = null;

        private async Task CheckForUpdates()
        {
            var graphClient = GetGraphClient();

            // get a page of messages
            var messages = await GetMessages(graphClient, DeltaLink);

            await GetAttachments(graphClient, messages);

            OutputMessages(messages);

            // go through all of the pages so that we can get the delta link on the last page.
            while (messages.NextPageRequest != null)
            {
                messages = messages.NextPageRequest.GetAsync().Result;
                OutputMessages(messages);
            }

            object deltaLink;

            if (messages.AdditionalData.TryGetValue("@odata.deltaLink", out deltaLink))
            {
                DeltaLink = deltaLink;
            }
        }

        private async Task GetAttachments(GraphServiceClient graphClient, IMessageDeltaCollectionPage messages)
        {
            foreach (var message in messages)
            {
                if (message.HasAttachments == true)
                {
                    IMessageAttachmentsCollectionPage attachmentsPage = await graphClient.Users["jleikam@integrativemeaning.com"]
                                        .MailFolders
                                        .Inbox
                                        .Messages[message.Id]
                                        .Attachments
                                        .Request()
                                        .GetAsync();

                   
                    if(attachmentsPage.CurrentPage.First().ODataType == "#microsoft.graph.fileAttachment")
                    {
                        var fileAttachment = attachmentsPage.CurrentPage.First() as FileAttachment;
                        Image image = Image.Load(fileAttachment.ContentBytes);
                        image.Save("testPic.jpg");
                    }
                   
                }
            }
        }

        private void OutputMessages(IMessageDeltaCollectionPage messages)
        {
            foreach (var message in messages)
            {
               
                var output = $"Message: {message.Id} {message.Subject}";
                
                Console.WriteLine(output);
            }
        }

        private async Task<IMessageDeltaCollectionPage> GetMessages(GraphServiceClient graphClient, object deltaLink)
        {
            IMessageDeltaCollectionPage page;

            if (lastPage == null)
            {
                page = await graphClient.Users["jleikam@integrativemeaning.com"]
                                        .MailFolders
                                        .Inbox
                                        .Messages
                                        .Delta()
                                        .Request()
                                        .GetAsync();
            }
            else
            {
                lastPage.InitializeNextPageRequest(graphClient, deltaLink.ToString());
                page = await lastPage.NextPageRequest.GetAsync();
            }

            lastPage = page;
            return page;
        }



    }
}