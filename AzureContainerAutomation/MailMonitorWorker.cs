using AutotaskPSA;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AzureContainerAutomation
{
    public class MailMonitorWorker : BackgroundService
    {
        private readonly IConfidentialClientApplication _confidentialClientApplication = null;
        private readonly GraphServiceClient _graphAPIClient = null;
        private readonly AutotaskAPIClient _atwsAPIClient = null;
        private readonly ILogger<MailMonitorWorker> _logger;
        private readonly MailboxConfig _configuration = null;
        private MailFolder _incoming_folder = null;
        private MailFolder _processed_folder = null;
        private MailFolder _failed_folder = null;
        private const int buffer_byteRateLimit = 200; // message bytes
        private const long _byteRateLimit = 10000000;
        private TimeSpan byteRateLimit_timer = new TimeSpan(0, 5, 0);
        private long current_byteRateLimit = 0;
        private DateTimeOffset _byteRateLimit_Reset_Time = DateTime.UtcNow;
        public MailMonitorWorker(ILogger<MailMonitorWorker> logger, MailboxConfig configuration)
        {

            if (logger == null || configuration == null)
            {
                throw new ArgumentNullException();
            }

            _logger = logger;
            _configuration = configuration;

            _logger.LogInformation("Worker({mailbox}) initializing at: {time}", _configuration.MailBox, DateTimeOffset.Now);
            try
            {
                _confidentialClientApplication = ConfidentialClientApplicationBuilder.Create(_configuration.Graph.ClientID)
                                              .WithClientSecret(_configuration.Graph.ClientSecret)
                                              .WithAuthority(new Uri(_configuration.Graph.ClientAuthority))
                                              .Build();
                _graphAPIClient = new GraphServiceClient(new MsalAuthenticationProvider(_confidentialClientApplication, new string[] { "https://graph.microsoft.com/.default" }));
                _atwsAPIClient = new AutotaskAPIClient(_configuration.Autotask.Username, _configuration.Autotask.Password, AutotaskIntegrationKey.nCentral);

                if (!SetupMailFolders())
                {
                    throw new NullReferenceException("Check graph permissions, connectivity, names. Unable to setup mail folders.");
                }

                _logger.LogInformation("Worker({mailbox}) initialized at: {time}", _configuration.MailBox, DateTimeOffset.Now);
            }
            catch (Exception _ex)
            {
                _logger.LogError(_ex, "Worker({mailbox}) failed initializing at: {time}", _configuration.MailBox, DateTimeOffset.Now);
                throw new InvalidOperationException();
            }
        }

        private bool SetupMailFolders()
        {
            bool _result = false;
            IUserMailFoldersCollectionPage _folderList = null;
            try
            {
                _folderList = _graphAPIClient.Users[_configuration.MailBox].MailFolders.Request().GetAsync().Result;
                while (_folderList != null)
                {
                    foreach (MailFolder _folder in _folderList)
                    {
                        if (_folder.DisplayName.ToLower() == _configuration.Folders.Incoming.ToLower())
                        {
                            _incoming_folder = _folder;
                        }

                        if (_folder.DisplayName.ToLower() == _configuration.Folders.Processed.ToLower())
                        {
                            _processed_folder = _folder;
                        }

                        if (_folder.DisplayName.ToLower() == _configuration.Folders.Failed.ToLower())
                        {
                            _failed_folder = _folder;
                        }
                    }
                    if (_folderList.NextPageRequest != null)
                    {
                        _folderList = _folderList.NextPageRequest.GetAsync().Result;
                    }
                    else
                        _folderList = null;
                }

                if (_incoming_folder == null)
                {
                    _incoming_folder = _graphAPIClient.Users[_configuration.MailBox].MailFolders.Request().AddAsync(new MailFolder()
                    {
                        DisplayName = _configuration.Folders.Incoming
                    }).Result;
                }
                if (_processed_folder == null)
                {
                    _processed_folder = _graphAPIClient.Users[_configuration.MailBox].MailFolders.Request().AddAsync(new MailFolder()
                    {
                        DisplayName = _configuration.Folders.Processed
                    }).Result;
                }
                if (_failed_folder == null)
                {
                    _failed_folder = _graphAPIClient.Users[_configuration.MailBox].MailFolders.Request().AddAsync(new MailFolder()
                    {
                        DisplayName = _configuration.Folders.Failed
                    }).Result;
                }
                if ((_incoming_folder != null) && (_processed_folder != null) && (_failed_folder != null))
                {
                    _result = true;
                }
            }
            catch
            {
                _result = false;
            }
            return _result;
        }
        protected override async System.Threading.Tasks.Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("Worker({mailbox}) started at: {time}", _configuration.MailBox, DateTimeOffset.Now);
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker({mailbox}) started mailbox processing at: {time}", _configuration.MailBox, DateTimeOffset.Now);
                string _filter = string.Empty;
                if (_configuration.Processing.UnreadOnly)
                {
                    _filter += $"isRead eq false";
                }

                if (IsValidEmail(_configuration.MailBox))
                {
                    _logger.LogInformation("Worker({mailbox}) processing mailbox: {mailbox} at: {time}", _configuration.MailBox, _configuration.MailBox, DateTimeOffset.Now);
                    try
                    {
                        IMailFolderMessagesCollectionPage _messages = await _graphAPIClient.Users[_configuration.MailBox].MailFolders[_incoming_folder.Id].Messages.Request().Header("Prefer", "outlook.body-content-type=\"text\"").Filter(_filter).GetAsync();
                        while (_messages != null)
                        {
                            foreach (Message _message in _messages)
                            {
                                await OnMessageReceivedAsync(_message);
                            }
                            if (_messages.NextPageRequest != null)
                                _messages = await _messages.NextPageRequest.Header("Prefer", "outlook.body-content-type=\"text\"").Filter(_filter).GetAsync();
                            else
                                _messages = null;
                        }
                    }
                    catch (Exception _ex)
                    {
                        _logger.LogError(_ex, "Worker({mailbox}) ERROR processing mailbox: {mailbox} at: {time}", _configuration.MailBox, _configuration.MailBox, DateTimeOffset.Now);
                    }
                }
                _logger.LogInformation("Worker({mailbox}) mailbox processing completed at: {time}", _configuration.MailBox, DateTimeOffset.Now);
                await System.Threading.Tasks.Task.Delay(_configuration.Processing.CheckDelay, stoppingToken);
            }
            _logger.LogInformation("Worker({mailbox}) Shutdown at: {time}", _configuration.MailBox, DateTimeOffset.Now);
        }
        private async System.Threading.Tasks.Task OnMessageReceivedAsync(Message message)
        {
            bool _successfully_processed = false;
            string _log_msgid = Sha256Sum(message.Id);
            _logger.LogInformation("Worker({mailbox}) processing message: {messageid} at: {time}", _configuration.MailBox, _log_msgid, DateTimeOffset.Now);

            string _ExistingTicketNumber = ExtractTicketNumber(message.Subject); // extract from subject, only first occurrence.
            Ticket _ExistingTicket = null;
            if (!string.IsNullOrWhiteSpace(_ExistingTicketNumber))
            {
                try
                {
                    _ExistingTicket = _atwsAPIClient.FindTicketByNumber(_ExistingTicketNumber);
                }
                catch
                {
                    _ExistingTicket = null;
                }
            }
            //Ticket and Ticket Note Creation
            if (_ExistingTicket == null)
            {
                Ticket _Created = await CreateAutoTaskTicketFromMessage(message);
                if (_Created != null)
                {
                    //_successfully_processed = await CreateAutoTaskAttachmentFromMesssageAttachmentsForTicket(_Created, message);
                    _successfully_processed = await CreateAutoTaskAttachmentFromMesssageForTicket(_Created, message);
                }
            }
            else
            {
                TicketNote _Created = await CreateAutoTaskTicketNoteFromMessage(_ExistingTicket, message);
                if (_Created != null)
                {
                    _successfully_processed = true;
                }
            }
            //Moving after processing/failures and mark as read.
            MailFolder _folder = _successfully_processed ? _processed_folder : _failed_folder;
            message = await MoveMessage(message, _folder);
            await MarkMessageRead(message);
        }

        private void byteRateLimiting(long sizeInBytes)
        {
            DateTimeOffset _now = DateTime.UtcNow;
            TimeSpan _time_Since = _now - _byteRateLimit_Reset_Time;
            TimeSpan _time_Remaining = byteRateLimit_timer - _time_Since;
            long _total_sizeInbytes = (sizeInBytes + buffer_byteRateLimit);
            if (_time_Since >= byteRateLimit_timer)
            {
                //we are >= 5 minutes we can reset
                current_byteRateLimit = 0;
                _byteRateLimit_Reset_Time = DateTime.UtcNow;
            }
            if ((current_byteRateLimit+_total_sizeInbytes) >= _byteRateLimit)
            {
                _logger.LogWarning("Worker rate limiting processing for {wait_time} at {time}", _time_Remaining, DateTime.Now);
                System.Threading.Thread.Sleep(_time_Remaining);
                byteRateLimiting(sizeInBytes);
            }
            else
            {
                current_byteRateLimit += _total_sizeInbytes;
            }
        }

        private async System.Threading.Tasks.Task<bool> CreateAutoTaskAttachmentFromMesssageForTicket(Ticket ticket, Message message)
        {
            string _log_msgid = Sha256Sum(message.Id);
            bool _result = false;
            if (ticket == null || message == null)
            {
                throw new ArgumentNullException();
            }
            try
            {
                _logger.LogInformation("Worker creating attachment from message: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
                HttpRequestMessage _request = _graphAPIClient.Users[_configuration.MailBox].Messages[message.Id].Request().GetHttpRequestMessage();
                _request.RequestUri = new Uri((_request.RequestUri.OriginalString + @"/$value"));
                byte[] _MIME_Message = Encoding.UTF8.GetBytes(await _graphAPIClient.HttpProvider.SendAsync(_request).Result.Content.ReadAsStringAsync());
                if (_MIME_Message != null)
                {
                    string _MIME_Name = $"{(string)ticket.TicketNumber}_Original.eml";
                    if (_MIME_Message.Length > 6000000) // zip it up
                    {
                        using (MemoryStream _memStream = new MemoryStream())
                        {
                            using (ZipArchive _archive = new ZipArchive(_memStream, ZipArchiveMode.Create, true))
                            {
                                ZipArchiveEntry _zipEntry = _archive.CreateEntry(_MIME_Name);
                                using (Stream _zipEntryStream = _zipEntry.Open())
                                {
                                    using (StreamWriter _zipStreamWriter = new StreamWriter(_zipEntryStream))
                                    {
                                        _zipStreamWriter.BaseStream.Write(_MIME_Message, 0, _MIME_Message.Length);
                                        _zipStreamWriter.Flush();
                                    }
                                }
                            }
                            _MIME_Message = _memStream.ToArray();
                            _MIME_Name += ".zip";
                        }
                    }
                    if (_MIME_Message.Length > 0 && _MIME_Message.Length <= 6000000)
                    {
                        try
                        {
                            byteRateLimiting(_MIME_Message.Length);
                            AutotaskPSA.Attachment _CreatedAttachment = _atwsAPIClient.CreateTicketAttachment(ticket, _MIME_Name, _MIME_Message, "All Autotask Users");
                            if (_CreatedAttachment != null)
                            {
                                _result = true;
                                _logger.LogInformation("Worker uploaded attachment of message: {messageid} at: {time} size: {size}", _log_msgid, DateTimeOffset.Now, _MIME_Message.Length);
                            }
                            else
                            {
                                throw new System.ServiceModel.CommunicationException("Failed to Create attachment, Networking, api username/password, upload limits, report to developer");
                            }
                        }
                        catch (Exception _ex)
                        {
                            _result = false;
                            _logger.LogError(_ex, "Worker failed uploading attachment of message: {messageid} at: {time} size: {size}", _log_msgid, DateTimeOffset.Now, _MIME_Message.Length);
                        }
                    }
                }
                else
                {
                    _result = true;
                }
            }
            catch (Exception _ex)
            {
                _result = false;
                _logger.LogError(_ex, "Worker failed creating attachment of message: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
            }
            return _result;
        }

        private async System.Threading.Tasks.Task MarkMessageRead(Message message)
        {
            string _log_msgid = Sha256Sum(message.Id);
            try
            {
                _logger.LogInformation("Worker marking message as read: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
                message.IsRead = true;
                await _graphAPIClient.Users[_configuration.MailBox].Messages[message.Id].Request().Select("IsRead").UpdateAsync(new Message { IsRead = true });
                _logger.LogInformation("Worker marked message as read: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
            }
            catch (Exception _ex)
            {
                _logger.LogError(_ex, "Worker failed to mark message as read : {messageid} at: {time} to: ", _log_msgid, DateTimeOffset.Now);
            }
        }

        private async Task<Message> MoveMessage(Message message, MailFolder _folder)
        {
            string _log_msgid = Sha256Sum(message.Id);
            Message _result = null;
            try
            {
                _logger.LogInformation("Worker moving message: {messageid} at: {time} to: ", _log_msgid, DateTimeOffset.Now, _folder.DisplayName);
                _result = await _graphAPIClient.Users[_configuration.MailBox].MailFolders[_incoming_folder.Id].Messages[message.Id].Move(_folder.Id).Request().PostAsync();
                _logger.LogInformation("Worker moved message: {messageid} at: {time} to: ", _log_msgid, DateTimeOffset.Now, _folder.DisplayName);
            }
            catch (Exception _ex)
            {
                _logger.LogError(_ex, "Worker failed to move message: {messageid} at: {time} to: ", _log_msgid, DateTimeOffset.Now, _folder.DisplayName);
            }
            return _result;
        }
#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        private async Task<Ticket> CreateAutoTaskTicketFromMessage(Message message)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            Ticket _result = null;
            string _log_msgid = Sha256Sum(message.Id);
            //EVERY TICKET MUST HAVE AN ACCOUNT BUT NOT A CONTACT, ATTEMPT TO FIND VIA CONTACT
            AutotaskPSA.Contact _MsgSender = null;
            Account _MsgSenderAccount = null;
            //lets see if we can find the sender.
            try
            {
                //we should really only get errors through these blocks if the connection fails or has bad data.
                _MsgSender = _atwsAPIClient.FindContactByEmail(message.Sender.EmailAddress.Address);
                if (_MsgSender != null)
                {
                    _MsgSenderAccount = _atwsAPIClient.FindAccountByID((int)_MsgSender.AccountID);
                }
                else
                {
                    _MsgSenderAccount = _atwsAPIClient.FindAccountByDomain(message.Sender.EmailAddress.Address.Split("@")[1]);
                }

                if (_MsgSenderAccount == null)
                {
                    _MsgSenderAccount = _atwsAPIClient.FindAccountByName(_configuration.Autotask.Defaults.Ticket.Account); // when all else fails use Defaults!
                    if (_MsgSenderAccount == null)
                    {
                        throw new ArgumentNullException("Unable to find Account, please check your default account settings.");
                    }
                }
            }
            catch
            {
                _MsgSender = null;
                _MsgSenderAccount = null;
            }
            //Create Ticket
            if (_MsgSenderAccount != null)
            {
                _logger.LogInformation("Worker creating new ticket for message: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
                try
                {
                    byteRateLimiting(message.Subject.Length+message.Body.Content.Length);
                    Ticket _CreatedTicket = _atwsAPIClient.CreateTicket(_MsgSenderAccount, _MsgSender, message.Subject, message.Body.Content, (DateTime.Now + _configuration.Autotask.Defaults.Ticket.DueDateOffset), _configuration.Autotask.Defaults.Ticket.Source, _configuration.Autotask.Defaults.Ticket.Status, _configuration.Autotask.Defaults.Ticket.Priority, _configuration.Autotask.Defaults.Ticket.Queue,_configuration.Autotask.Defaults.Ticket.WorkType);
                    if (_CreatedTicket != null && !string.IsNullOrWhiteSpace((string)_CreatedTicket.TicketNumber))
                    {
                        _result = _CreatedTicket;
                        _logger.LogInformation("Worker created ticket: {ticketnumber} for message: {messageid} at: {time}", _CreatedTicket.TicketNumber, _log_msgid, DateTimeOffset.Now);
                        //handle attachments.
                    }
                    else
                    {
                        throw new ApplicationException("Ticket Creation failed check defaults, contact developer");
                    }
                }
                catch (Exception _ex)
                {
                    _logger.LogError(_ex, "Worker failed creating new ticket for message: {messageid} at: {time}", _log_msgid, DateTimeOffset.Now);
                }
            }
            return _result;
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        private async Task<TicketNote> CreateAutoTaskTicketNoteFromMessage(Ticket _ExistingTicket, Message message)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            TicketNote _result = null;
            string _log_msgid = Sha256Sum(message.Id);
            Resource _MsgSender = null;
            //Create Ticket Note
            //Resource impersonation isn't a requirement. don't try too hard to find a resource. fail gracefully.
            try
            {
                _MsgSender = _atwsAPIClient.FindResourceByEmail(message.Sender.EmailAddress.Address);
            }
            catch
            {
                _MsgSender = null; //graceful failures are the best failures.
            }
            _logger.LogInformation("Worker adding note to ticket: {ticketnumber} for message: {messageid} at: {time}", (string)_ExistingTicket.TicketNumber, _log_msgid, DateTimeOffset.Now);
            try
            {
                byteRateLimiting(message.Subject.Length + message.Body.Content.Length);
                TicketNote _CreatedTicketNote = _atwsAPIClient.CreateTicketNote(_ExistingTicket, _MsgSender, message.Subject.Replace((string)_ExistingTicket.TicketNumber, string.Empty), message.Body.Content, _configuration.Autotask.Defaults.TicketNote.Type, _configuration.Autotask.Defaults.TicketNote.Publish);
                if (_CreatedTicketNote != null)
                {
                    _result = _CreatedTicketNote;
                    _logger.LogInformation("Worker added ticket note to ticket : {ticketnumber} for message: {messageid} at: {time}", (string)_ExistingTicket.TicketNumber, _log_msgid, DateTimeOffset.Now);
                }
            }
            catch (Exception _ex)
            {
                _logger.LogError(_ex, "Worker failed adding ticket note to ticket : {ticketnumber} for message: {messageid} at: {time}", (string)_ExistingTicket.TicketNumber, _log_msgid, DateTimeOffset.Now);
            }
            return _result;
        }

        protected internal string Sha256Sum(string rawData)
        {
            // Create a SHA256   
            using (SHA256 sha256Hash = SHA256.Create())
            {
                // ComputeHash - returns byte array  
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(rawData));

                // Convert byte array to a string   
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }
        protected internal bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            try
            {
                // Normalize the domain
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));

                // Examines the domain part of the email and normalizes it.
                string DomainMapper(Match match)
                {
                    // Use IdnMapping class to convert Unicode domain names.
                    IdnMapping idn = new IdnMapping();

                    // Pull out and process domain name (throws ArgumentException on invalid)
                    string domainName = idn.GetAscii(match.Groups[2].Value);

                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email,
                    @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
        }
        protected internal string ExtractTicketNumber(string data)
        {
            const string _AT_TicketRegEx = "[T][0-9]{8}[.][0-9]{4}";
            string _searchResult = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(data))
                {
                    MatchCollection _matches = Regex.Matches(data, _AT_TicketRegEx, RegexOptions.Multiline);
                    if (_matches.Count > 0)
                    {
                        _searchResult = _matches.First().Value;
                    }
                }
                return _searchResult;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
