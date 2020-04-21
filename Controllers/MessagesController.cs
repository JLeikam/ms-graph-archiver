using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ms_graph_app.Models;
using Newtonsoft.Json;
using System.Threading;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Microsoft.Azure.CognitiveServices.Vision.ComputerVision;
using Microsoft.Azure.CognitiveServices.Vision.ComputerVision.Models;
using CsvHelper;
using System.Text;
using System.Net.Http;
using ms_graph_app.Utils;

namespace ms_graph_app.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MessagesController : ControllerBase
    {
        private readonly GraphConfig config;
        private GraphHelper graphHelper;

        public MessagesController(GraphHelper graphHelper)
        {
            this.graphHelper = graphHelper;
            this.config = graphHelper.config;
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

        private async Task CheckForUpdates()
        {
            var graphClient = graphHelper.GetGraphClient();

            IMailFolderMessagesCollectionPage messages = await GetUnreadMessages(graphClient);

            var msgToAttachmentsDict = await GetFileAttachments(graphClient, messages);


            foreach (KeyValuePair<Message, List<FileAttachment>> kvp in msgToAttachmentsDict)
            {
                foreach (FileAttachment attachment in kvp.Value)
                {
                    if (attachment.ContentType.Contains("image"))
                    {
                        string txtFromImg = await OCR(attachment.ContentBytes);
                        var dataName = "name:fileBlock1";
                        var msgTitle = kvp.Key.Subject;
                        //var pageTitle = String.Join(" ", txtFromImg.Split().Take(5).ToArray());
                        var htmlString =
                            "<!DOCTYPE html>" +
                            "<html>" +
                            "<head>" +
                            $"<title> {msgTitle} </title>" +
                            "</head>" +
                            "<body>";

                        htmlString += $"<p>{txtFromImg}</p>";
                     
                        htmlString += $"<object data-attachment=\"{attachment.Name}\" data=\"{dataName}\" type=\"{attachment.ContentType}\" />";

                        htmlString += "</body>"
                                    + "</html>";
                        Console.WriteLine(htmlString);
                        await PostToNotebook(graphClient, htmlString, attachment);
                    }
                    else if (attachment.ContentType.Contains("csv"))
                    {
                        var dataName = "name:fileBlock1";
                        var records = CsvParse(attachment.ContentBytes);
                        var msgTitle = kvp.Key.Subject;
                        var htmlString =
                            "<!DOCTYPE html>" +
                            "<html>" +
                            "<head>" +
                            $"<title> {msgTitle} </title>" +
                            "</head>" +
                            "<body>";
                        foreach (var record in records)
                        {
                            htmlString += $"<p>{record.Annotation} ({record.Location})</p>";
                        }
                        htmlString += $"<object data-attachment=\"{attachment.Name}\" data=\"{dataName}\" type=\"{attachment.ContentType}\" />";
                        htmlString += "</body>"
                                    + "</html>";
                        Console.WriteLine(htmlString);
                        await PostToNotebook(graphClient, htmlString, attachment);
                    }
                }
                await MarkMessageAsRead(graphClient, kvp.Key.Id);
            }

            OutputMessages(messages);

        }

        private async Task PostToNotebook(GraphServiceClient graphClient, string msg, FileAttachment attachment)
        {
            
            var accessToken = graphHelper.GetAccessToken().Result;
            using (var client = new HttpClient())
            {

                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                using (var content = new MultipartFormDataContent("MyPartBoundary198374"))
                {
                    var stringContent = new StringContent(msg, Encoding.UTF8, "text/html");
                    content.Add(stringContent, "Presentation");
                    var fileContent = new ByteArrayContent(attachment.ContentBytes);
                    fileContent.Headers.ContentType = new MediaTypeHeaderValue(attachment.ContentType);
                    content.Add(fileContent, "fileBlock1", "fileBlock1");
                    var requestUrl = graphClient.Users["jleikam@integrativemeaning.com"]
                        .Onenote
                        .Pages
                        .RequestUrl;
                    using (
                       var message =
                           await client.PostAsync(requestUrl, content))
                    {
                        Console.WriteLine(message.ReasonPhrase);
                    }
                }
            }
        }

        private async Task MarkMessageAsRead(GraphServiceClient graphClient, string msgId)
        {
            var msg = await graphClient.Users["jleikam@integrativemeaning.com"]
                .Messages[msgId]
                .Request()
                .Select("IsRead")
                .UpdateAsync(new Message()
                {
                    IsRead = true
                });
        }

        private async Task<Dictionary<Message, List<FileAttachment>>> GetFileAttachments(GraphServiceClient graphClient, IMailFolderMessagesCollectionPage messages)
        {
            var msgToAttachmentsDict = new Dictionary<Message, List<FileAttachment>>();
            for (int i = 0; i < messages.Count; i++)
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

                    for (int j = 0; j < attachmentsPage.Count; j++)
                    {
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

        private List<KindleCsv> CsvParse(byte[] contentBytes)
        {
            var records = new List<KindleCsv>();
            Stream stream = new MemoryStream(contentBytes);
            using (var reader = new StreamReader(stream))
            using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
            {
                

                //skip the first 7 lines because they contain irrelevant information
                for(int i =1; i<=7; i++)
                {
                    csv.Read();
                }
                csv.Read();
                csv.ReadHeader();
                while (csv.Read())
                {
                    
                    var record = new KindleCsv
                    {
                        AnnotationType = csv.GetField("Annotation Type"),
                        Location = csv.GetField("Location"),
                        IsStarred = csv.GetField("Starred?"),
                        Annotation = csv.GetField("Annotation")
                    };
                    records.Add(record);
                    Console.WriteLine(record.Annotation);
                }
            }
            Console.WriteLine("END OF PARSE");
            return records;
        }

        private ComputerVisionClient GetComputerVisionClient()
        {
            ComputerVisionClient client = new ComputerVisionClient(new ApiKeyServiceClientCredentials(config.SubscriptionKey))
            { Endpoint = config.Endpoint };

            return client;
        }

        private async Task<string> OCR(byte[] contentBytes)
        {

            var textFromImage = "";
            var visionClient = GetComputerVisionClient();
            Stream stream = new MemoryStream(contentBytes);
            BatchReadFileInStreamHeaders textHeaders = await visionClient.BatchReadFileInStreamAsync(stream);
            string operationLocation = textHeaders.OperationLocation;
            const int numberOfCharsInOperationId = 36;
            string operationId = operationLocation.Substring(operationLocation.Length - numberOfCharsInOperationId);

            // Extract the text
            // Delay is between iterations and tries a maximum of 10 times.
            int i = 0;
            int maxRetries = 10;
            ReadOperationResult results;
            Console.WriteLine($"Extracting text from image");
            Console.WriteLine();

            do
            {
                results = await visionClient.GetReadOperationResultAsync(operationId);
                Console.WriteLine("Server status: {0}, waiting {1} seconds...", results.Status, i);
                await Task.Delay(1000);
                if (i == 9)
                {
                    Console.WriteLine("Server timed out.");
                }
            }
            while ((results.Status == TextOperationStatusCodes.Running || results.Status == TextOperationStatusCodes.NotStarted) && i++ < maxRetries);

            // Display the found text.
            Console.WriteLine();
            var textRecognitionLocalFileResults = results.RecognitionResults;
            foreach (TextRecognitionResult recResult in textRecognitionLocalFileResults)
            {
                foreach (Line line in recResult.Lines)
                {
                    textFromImage += " " + line.Text;
                    Console.WriteLine(line.Text);
                }
            }
            Console.WriteLine();

            return textFromImage;
        }
    }
}