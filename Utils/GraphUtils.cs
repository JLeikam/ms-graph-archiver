using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace ms_graph_app.Utils
{
    public class GraphHelper
    {
        public readonly GraphConfig config;
        private static Dictionary<string, Subscription> Subscriptions = new Dictionary<string, Subscription>();
        private static Timer subscriptionTimer = null;
        private static int MAX_EXPIRATION_MINUTES = 4230;
        public GraphHelper(GraphConfig config)
        {
            this.config = config;
        }
        public GraphServiceClient GetGraphClient()
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

        public async Task<string> GetAccessToken()
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

        public async Task InitSubscription()
        {
            var graphServiceClient = GetGraphClient();
            var sub = new Subscription
            {
                ChangeType = "created",
                NotificationUrl = $"{config.URL}/api/messages",
                Resource = $"/users/jleikam@integrativemeaning.com/mailFolders/{config.ArchiverId}/messages",
                ExpirationDateTime = DateTime.UtcNow.AddMinutes(MAX_EXPIRATION_MINUTES),
                ClientState = Guid.NewGuid().ToString()
            };

            var newSubscription = await graphServiceClient
                .Subscriptions
                .Request()
                .AddAsync(sub);

            Subscriptions[newSubscription.Id] = newSubscription;

            if (subscriptionTimer == null)
            {
                //check subscriptions every hour
                subscriptionTimer = new Timer(CheckSubscriptions, null, 5000, 3600000);
            }

            Console.WriteLine($"Subscribed. Id: {newSubscription.Id}, Expiration: {newSubscription.ExpirationDateTime}");
        }

        public async void RenewSubscription(Subscription subscription)
        {
            Console.WriteLine($"Current subscription: {subscription.Id}, Expiration: {subscription.ExpirationDateTime}");

            var graphServiceClient = GetGraphClient();

            var newSubscription = new Subscription
            {
                ExpirationDateTime = DateTime.UtcNow.AddMinutes(MAX_EXPIRATION_MINUTES)
            };

            await graphServiceClient
              .Subscriptions[subscription.Id]
              .Request()
              .UpdateAsync(newSubscription);

            subscription.ExpirationDateTime = newSubscription.ExpirationDateTime;
            Console.WriteLine($"Renewed subscription: {subscription.Id}, New Expiration: {subscription.ExpirationDateTime}");
        }

        private void CheckSubscriptions(Object stateInfo)
        {
            AutoResetEvent autoEvent = (AutoResetEvent)stateInfo;

            Console.WriteLine($"Checking subscriptions {DateTime.Now.ToString("h:mm:ss.fff")}");
            Console.WriteLine($"Current subscription count {Subscriptions.Count()}");

            foreach (var subscription in Subscriptions)
            {
                // if the subscription expires in the next two hours, renew it
                if (subscription.Value.ExpirationDateTime < DateTime.UtcNow.AddMinutes(120))
                {
                    RenewSubscription(subscription.Value);
                }
            }
        }
    }
}
