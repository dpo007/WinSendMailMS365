using Azure.Identity;
using Microsoft.Graph;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace WinSendMailMS365
{
    internal class WinSendMailMS365
    {
        private static Stream GenerateStreamFromString(string s)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        private static async Task Main(string[] args)
        {
            string rawEmail = null;
            string line;

            // Read from console/stdin until "Ctrl-Z"...
            while ((line = Console.ReadLine()) != null)
            {
                rawEmail += line + Environment.NewLine;
            }

            if (Properties.Settings.Default.SaveEmailsToDisk)
            {
                // Show raw email in console.
                Console.Write(rawEmail);

                // Save the email to a file on disk.
                string rndFileNamePart = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                StreamWriter streamWriter = new StreamWriter(@"WinSendMailLog-" + rndFileNamePart + ".txt");
                streamWriter.Write(rawEmail);
                streamWriter.Dispose();
            }

            // Import raw email into a MimeMessage object.
            _ = new MimeMessage();
            MimeMessage mimeEmailMsg = MimeMessage.Load(GenerateStreamFromString(rawEmail));

            // Load MS365 App Integration info.
            string tenantId = Properties.Settings.Default.MS365TenantID;
            string clientId = Properties.Settings.Default.MS365ClientID;
            string clientSecret = Properties.Settings.Default.MS365ClientSecret;

            // using Azure.Identity;
            TokenCredentialOptions options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            GraphServiceClient graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            // Define email to send via Graph API.
            Recipient sender = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = mimeEmailMsg.Sender.ToString()
                }
            };

            Recipient from = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = mimeEmailMsg.From[0].ToString()
                }
            };

            List<Recipient> recipients = new List<Recipient>();
            foreach (InternetAddress address in mimeEmailMsg.To)
            {
                Recipient recipient = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = address.ToString()
                    }
                };

                recipients.Add(recipient);
            }

            // Prep email Body.
            string preppedBodyContent;
            if (Properties.Settings.Default.HTMLDecodeContent)
            {
                // Decode HTML content.
                preppedBodyContent = WebUtility.HtmlDecode(mimeEmailMsg.GetTextBody(MimeKit.Text.TextFormat.Plain));
            }
            else
            {
                // Do not decode HTML content.
                preppedBodyContent = mimeEmailMsg.GetTextBody(MimeKit.Text.TextFormat.Plain);
            }

            // Prep email Subject.
            string preppedSubject;
            if (Properties.Settings.Default.HTMLDecodeContent)
            {
                // Decode HTML content.
                preppedSubject = WebUtility.HtmlDecode(mimeEmailMsg.Subject);
            }
            else
            {
                // Do not decode HTML content.
                preppedSubject = mimeEmailMsg.Subject;
            }

            // Assemble email message.
            Message message = new Message
            {
                Subject = preppedSubject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = preppedBodyContent
                },
                ToRecipients = recipients,
                Sender = sender,
                From = from
            };

            // Get ID of user who we'll send mail as.
            string sendingUserUPN = Properties.Settings.Default.MS365UserName;
            User user = await graphClient
                .Users[sendingUserUPN]
                .Request()
                .GetAsync();

            // Send the email.
            await graphClient.Users[user.Id]
                .SendMail(message, false)
                .Request()
                .PostAsync();
        }
    }
}