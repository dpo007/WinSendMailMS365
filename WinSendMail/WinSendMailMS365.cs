﻿using Azure.Identity;
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
            using (MemoryStream stream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.Write(s);
                    writer.Flush();
                }

                stream.Position = 0;
                return stream;
            }
        }

        private static async Task Main()
        {
            string rawEmail = null;
            string line;

            // Read from console/stdin until "Ctrl-Z"...
            while ((line = Console.ReadLine()) != null)
            {
                rawEmail += $"{line}{Environment.NewLine}";
            }

            if (Properties.Settings.Default.SaveEmailsToDisk)
            {
                // Show raw email in console.
                Console.Write(rawEmail);

                // Save the email to a file on disk.
                string rndFileNamePart = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                StreamWriter streamWriter = new StreamWriter($"RawInput-{DateTime.Now:yyyyMMdd_HHmmss}-{rndFileNamePart}.txt");
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
            TokenCredentialOptions credentialOptions = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, credentialOptions);

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            GraphServiceClient graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            // Define email to send via Graph API.
            Recipient emailSender = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = mimeEmailMsg.Sender.ToString()
                }
            };

            Recipient emailFrom = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = mimeEmailMsg.From[0].ToString()
                }
            };

            List<Recipient> emailRecipients = new List<Recipient>();
            foreach (InternetAddress address in mimeEmailMsg.To)
            {
                Recipient recipient = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = address.ToString()
                    }
                };

                emailRecipients.Add(recipient);
            }

            // Prep email Body.
            string emailBodyContent;
            if (Properties.Settings.Default.HTMLDecodeContent)
            {
                // Decode HTML content.
                emailBodyContent = WebUtility.HtmlDecode(mimeEmailMsg.GetTextBody(MimeKit.Text.TextFormat.Plain));
            }
            else
            {
                // Do not decode HTML content.
                emailBodyContent = mimeEmailMsg.GetTextBody(MimeKit.Text.TextFormat.Plain);
            }

            // Prep email Subject.
            string emailSubject;
            if (Properties.Settings.Default.HTMLDecodeContent)
            {
                // Decode HTML content.
                emailSubject = WebUtility.HtmlDecode(mimeEmailMsg.Subject);
            }
            else
            {
                // Do not decode HTML content.
                emailSubject = mimeEmailMsg.Subject;
            }

            // Assemble email message.
            Message emailMessage = new Message
            {
                Subject = emailSubject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = emailBodyContent
                },
                ToRecipients = emailRecipients,
                Sender = emailSender,
                From = emailFrom
            };

            // Attempt to look up User via Graph API.
            string sendingUserUPN = Properties.Settings.Default.MS365UserName;
            User sendingUser;
            try
            {
                // Get ID of user who we'll send mail as.
                sendingUser = await graphClient
                    .Users[sendingUserUPN]
                    .Request()
                    .GetAsync();
            }
            catch (Exception ex)
            {
                // Log error to file.
                LogError($"Problem looking up user with UPN \"{sendingUserUPN}\".  Error: {ex.Message}");
                throw;
            }

            // Attempt to send email.
            if (sendingUser == null)
            {
                // Log error to file.
                string errMsg = $"User not found.  UPN: {sendingUserUPN}";
                LogError(errMsg);
                throw new Exception(errMsg);
            }
            else
            {
                try
                {
                    // Send email.
                    await graphClient
                        .Users[sendingUser.Id]
                        .SendMail(emailMessage, true)
                        .Request()
                        .PostAsync();
                }
                catch (Exception ex)
                {
                    // Log error to file.
                    LogError($"Problem sending email.  Error: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Logs an error message to a file.
        /// </summary>
        /// <param name="errorMessage">Error message to log.</param>
        private static void LogError(string errorMessage)
        {
            // Build filename from today's date and time.
            string sendErrorLogFileName = $"SendError-{DateTime.Now:yyyy-MM-dd}.log";

            // Create log file if it doesn't exist.
            if (!System.IO.File.Exists(sendErrorLogFileName))
            {
                System.IO.File.Create(sendErrorLogFileName).Dispose();
            }

            // Prefix error message with date/time.
            errorMessage = $"[{DateTime.Now:yyyy-MM-dd hh:mm:ss.fff tt}] :: {errorMessage}";

            // Append error message to log file.
            System.IO.File.AppendAllText(sendErrorLogFileName, $"{errorMessage}{Environment.NewLine}");
        }
    }
}