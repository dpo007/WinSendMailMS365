﻿using Azure.Identity;
using Microsoft.Graph;
using MimeKit;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace WinSendMailMS365
{
    internal class WinSendMailMS365
    {
        private static async Task Main()
        {
            string rawEmail = null;
            string line;

            // Load application settings from file.
            AppSettings appSettings = LoadAppSettings();

            // Read from console/stdin until "Ctrl-Z"...
            while ((line = Console.ReadLine()) != null)
            {
                rawEmail += $"{line}{Environment.NewLine}";
            }

            if (appSettings.SaveEmailsToDisk)
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
            MimeMessage mimeEmailMsg = GenerateMimeMsgFromString(rawEmail);

            // Set options to use for Azure.Identity credential.
            TokenCredentialOptions credentialOptions = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // Create Azure.Identity 'Client Secret' credential (based on application settings).
            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(
                appSettings.MS365TenantID,
                appSettings.MS365ClientID,
                appSettings.MS365ClientSecret,
                credentialOptions);

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            GraphServiceClient graphClient = new GraphServiceClient(
                clientSecretCredential,
                scopes);

            // Create a new Graph message to send.
            Message emailMessage = CreateGraphEmailMessage(mimeEmailMsg, appSettings.HTMLDecodeContent);

            // Attempt to look up User via Graph API.
            string sendingUserUPN = appSettings.MS365SendingUser;
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
        /// Create a MimeMessage from a raw email string.
        /// </summary>
        /// <param name="emailString">Raw email string to generate MimeMessage from.</param>
        /// <returns>MimeMessage containing contents of input string.</returns>
        private static MimeMessage GenerateMimeMsgFromString(string emailString)
        {
            MimeMessage mimeEmailMsg = new MimeMessage();

            using (MemoryStream stream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.Write(emailString);
                    writer.Flush();
                    stream.Position = 0;
                    mimeEmailMsg = MimeMessage.Load(stream);
                }
            }

            return mimeEmailMsg;
        }

        /// <summary>
        /// Create Graph email message to send via API.
        /// </summary>
        /// <param name="mimeEmailMsg">MimeMessage email object.</param>
        /// <param name="decodeHTML">Decode HTML in Subject and Body?</param>
        /// <returns>Crafted Graph email Message object</returns>
        private static Message CreateGraphEmailMessage(MimeMessage mimeEmailMsg, bool decodeHTML)
        {
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
            if (decodeHTML)
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
            if (decodeHTML)
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

            return emailMessage;
        }

        /// <summary>
        /// Load AppSettings from AppSettings.json file.
        /// </summary>
        /// <returns>Object with application settings.</returns>
        private static AppSettings LoadAppSettings()
        {
            // Create new AppSettings object.
            AppSettings appSettings = new AppSettings
            {
                // Populate AppSettings object with default settings.
                MS365TenantID = "Your TenantID",
                MS365ClientID = "App ClientID",
                MS365ClientSecret = "App ClientSecret",
                MS365SendingUser = "sendingUser@yourcompany.com",
                SaveEmailsToDisk = false,
                HTMLDecodeContent = true
            };

            // Check if AppSettings.json JSON file exists, create new file if it does not.
            if (!System.IO.File.Exists("AppSettings.json"))
            {
                System.IO.File.Create("AppSettings.json").Dispose();

                // Serialize AppSettings to JSON, save to new config file.
                string json = JsonConvert.SerializeObject(appSettings, Formatting.Indented);
                System.IO.File.WriteAllText("AppSettings.json", json);

                LogError("AppSettings.json file not found.  Created new file with default settings.  Please edit it and try again.");

#if DEBUG
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
#endif
                // Exit program.
                Environment.Exit(0);
            }

            // Load settings by deserializing AppSettings.json into an AppSettings object.
            string jsonConfig = System.IO.File.ReadAllText("AppSettings.json");
            appSettings = JsonConvert.DeserializeObject<AppSettings>(jsonConfig);

            // Return AppSettings object.
            return appSettings;
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

            // Write error to console.
            Console.WriteLine(errorMessage);

            // Append error to log file.
            System.IO.File.AppendAllText(sendErrorLogFileName, $"{errorMessage}{Environment.NewLine}");
        }
    }
}