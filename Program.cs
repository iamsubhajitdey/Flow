namespace Flow
{
    using System;
    using System.Collections.Generic;
    using Razorpay.Api;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Gmail.v1;
    using Google.Apis.Gmail.v1.Data;
    using Google.Apis.Services;
    using Google.Apis.Util.Store;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Net.Mail;

    using MailKit.Security;
    using MailKit.Net.Smtp;
    using MailKit;
    using MimeKit;
    using MimeKit.Text;
    using Microsoft.IdentityModel.Tokens;

    class Program
    {
        public static string senderEmail = "placeholder@email.com";
        public static string senderName = "placeholder";
        public static string productZipFolder = "//placeholder//path.zip";
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Flow";




        static void Main(string[] args)
        {
            //--------------RAZORPAY INTEGRATION-------------------------
            string[] RazorpayDetails = File.ReadAllLines("razorpaycredentials.txt");
            RazorpayClient client = new RazorpayClient(RazorpayDetails[0], RazorpayDetails[1]);

            Dictionary<string, object> options = new Dictionary<string, object>();

            long existingTimestamp = Convert.ToInt64(File.ReadAllText("timestamp.txt"));
            List<Payment> payments = client.Payment.All(options);
            var listPayments = payments
                                .Where(p => p["created_at"].Value > existingTimestamp)
                                .Where(p => p["status"].Value == "captured");
            var failedPayments = payments.Where(p => p["status"].Value == "failed");


            //----------------GMAIL Integration-------------------------------
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //Define parameters of request.
            List<Message> messages = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
            request.LabelIds = new string[] { "SENT" };
            ListMessagesResponse response = request.Execute();
            var mostRecentSentMessageId = response.Messages.First().Id;
            var mostRecentSentMessage = service.Users.Messages.Get("me", mostRecentSentMessageId).Execute();
            var toAddress = mostRecentSentMessage.Payload.Headers.Where(h => h.Name == "To").First().Value;

            //Remove any payments that have been emailed -- CHECK FOR MULTIPLE ORDERS FROM SAME EMAIL
            List<Payment> newPayments = new List<Payment>();
            foreach (var payment in listPayments)
            {
                string email = payment["email"].Value;
                if (email == toAddress)
                {
                    break;
                }

                newPayments.Add(payment);
            }

            //Going through payments to process
            foreach (var payment in newPayments)
            {
                long amount = payment["amount"].Value;
                string email = payment["email"].Value;
                bool captured = payment["captured"].Value;
                string id = payment["id"].Value;
                var notes = payment["notes"];
                string product = (notes["product"] == null) ? string.Empty : notes["product"];
                long timestamp = payment["created_at"].Value;

                if (timestamp != 0)
                {

                    if (timestamp > existingTimestamp)
                    {
                        existingTimestamp = timestamp;
                    }
                }

                if (string.IsNullOrWhiteSpace(product))
                {
                    Console.WriteLine("Transaction without product value present : email is " + email);
                    File.WriteAllText("timestamp.txt", existingTimestamp.ToString());
                    return;
                }

                MimeMessage emailContent = createEmail(email, "", product);
                Message message = createMessageWithEmail(emailContent);
                var result = service.Users.Messages.Send(message, "me").Execute();
            }

            File.WriteAllText("timestamp.txt", existingTimestamp.ToString());

        }

        public static MimeMessage createEmail(string to, string from, string product)
        {
            string path = productZipFolder;
            MailMessage mail = new MailMessage();
            mail.Subject = "Placeholder : Thanks for your support!";
            mail.Sender = new MailAddress(senderEmail, senderName);

            MimeKit.MimeMessage messageToSend = MimeKit.MimeMessage.CreateFromMailMessage(mail);

            messageToSend.To.Add(new MailboxAddress(to));
            var body = new TextPart(TextFormat.Text)
            {
                Text = "Hi ☺\n\n" +
            "Thanks for your support and hope you enjoy what we've created! \n\n" +
            "Download and unzip the contents attached to find your purchase.\n\n" +
            "Please respond to this email if you have any questions or face any issues. \n\n" +
            "Thanking You,\n" +
            "- Placeholder"
            };

            var attachment = new MimePart("image", "gif")
            {
                Content = new MimeContent(File.OpenRead(path), ContentEncoding.Default),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = Path.GetFileName(path)
            };

            var multipart = new Multipart("mixed");
            multipart.Add(body);
            multipart.Add(attachment);
            messageToSend.Body = multipart;

            return messageToSend;

        }

        public static Message createMessageWithEmail(MimeMessage emailContent)
        {
            var memory = new MemoryStream();
            emailContent.WriteTo(memory);

            string encodedEmail = Base64UrlEncoder.Encode(memory.ToArray());
            Message message = new Message();
            message.Raw = encodedEmail;
            return message;
        }
    }
}
