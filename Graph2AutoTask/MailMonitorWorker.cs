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
using Graph2AutoTask.ApiQueue;
using System.Collections.Generic;
using OpsGenieApi;
using SkiaSharp;

namespace Graph2AutoTask
{
    public class MailMonitorWorker : BackgroundService
    {
        private readonly IConfidentialClientApplication _confidentialClientApplication = null;
        private readonly GraphServiceClient _graphAPIClient = null;
        private readonly AutotaskAPIClient _atwsAPIClient = null;
        private readonly ILogger<MailMonitorWorker> _logger;
        private readonly MailboxConfig _configuration = null;
        private readonly OpsGenieClient _opsGenieClient = null;
        private ApiQueueManager _queueManager = null;
        private MimeTypeDatabase _mimeDatabase = new MimeTypeDatabase();
        private MailFolder _incoming_folder = null;
        private MailFolder _processed_folder = null;
        private MailFolder _failed_folder = null;

        public MailMonitorWorker(ILogger<MailMonitorWorker> logger, MailboxConfig configuration)
        {

            if (logger == null || configuration == null)
            {
                throw new ArgumentNullException();
            }

            _logger = logger;
            _configuration = configuration;
            _logger.LogInformation($"[{_configuration.MailBox}] - initializing");
            try
            {
                _confidentialClientApplication = ConfidentialClientApplicationBuilder.Create(_configuration.Graph.ClientID)
                                              .WithClientSecret(_configuration.Graph.ClientSecret)
                                              .WithAuthority(new Uri(_configuration.Graph.ClientAuthority))
                                              .Build();
                _opsGenieClient = new OpsGenieClient(new OpsGenieClientConfig() { ApiKey = _configuration.OpsGenie.ApiKey }, new OpsGenieSerializer());
                _graphAPIClient = new GraphServiceClient(new MsalAuthenticationProvider(_confidentialClientApplication, new string[] { "https://graph.microsoft.com/.default" }));
                _atwsAPIClient = new AutotaskAPIClient(_configuration.Autotask.Username, _configuration.Autotask.Password, AutotaskIntegrationKey.nCentral);

                if (!SetupMailFolders())
                {
                    throw new NullReferenceException("Check graph permissions, connectivity, names. Unable to setup mail folders.");
                }

                _logger.LogInformation($"[{_configuration.MailBox}] - initialized");
            }
            catch (Exception _ex)
            {
                _logger.LogError(_ex, $"[{_configuration.MailBox}] - failed initializing");
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
            _logger.LogInformation($"[{_configuration.MailBox}] - started execution");
            _queueManager = new ApiQueueManager(stoppingToken, _configuration, _logger, _opsGenieClient);
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - started mailbox processing");
                string _filter = string.Empty;
                if (_configuration.Processing.UnreadOnly)
                {
                    _filter += $"isRead eq false";
                }

                if (IsValidEmail(_configuration.MailBox))
                {
                    _logger.LogInformation($"[{_configuration.MailBox}] - processing mailbox");
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
                        _logger.LogError(_ex, $"[{_configuration.MailBox}] - ERROR processing mailbox");
                    }
                }
                _logger.LogInformation($"[{_configuration.MailBox}] - mailbox processing completed");
                await System.Threading.Tasks.Task.Delay(_configuration.Processing.CheckDelay, stoppingToken);
            }
            _logger.LogInformation($"[{_configuration.MailBox}] - shutdown");
        }
        private void queue_FindTicketByNumber(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            Ticket _ExistingTicket = null;
            string _ExistingTicketNumber = ExtractTicketNumber(_internal.Message.Subject);
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindTicketByNumber({_internal.ID})");
            if (!string.IsNullOrWhiteSpace(_ExistingTicketNumber))
            {
                try
                {
                    _ExistingTicket = _atwsAPIClient.FindTicketByNumber(_ExistingTicketNumber);
                }
                catch(Exception _ex)
                {
                    Exception _included = _ex;
                    if (_ex.InnerException != null)
                        _included = _ex.InnerException;
                    _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindTicketByNumber({_internal.ID})");
                    throw new ApiQueueException($"queue_FindTicketByNumber({_internal.ID})", _included);
                }
            }
            if (_ExistingTicket != null)
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindTicketByNumber({_internal.ID})");
                //create a ticket note by first seeing if the sender is a resource!
                _arguments.Add("ticket", _ExistingTicket);
                ApiQueueJob _findResourceByEmail = new ApiQueueJob(_internal.ID, queue_FindResourceByEmail, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findResourceByEmail);
            }
            else
            {
                //create a ticket - start by finding a sender account by email
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindTicketByNumber({_internal.ID})");
                ApiQueueJob _findContactByEmail = new ApiQueueJob(_internal.ID, queue_FindContactByEmail, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findContactByEmail);
            }
        }
        private void queue_FindContactByEmail(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Contact _ExistingContact = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindContactByEmail({_internal.ID})");
            try
            {
                _ExistingContact = _atwsAPIClient.FindContactByEmail(_internal.Message.Sender.EmailAddress.Address);
            }
            catch(Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindContactByEmail({_internal.ID})");
                throw new ApiQueueException($"queue_FindContactByEmail({_internal.ID})", _included);
            }
            if (_ExistingContact != null)
            {
                //we have a contact - find account from contact
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindContactByEmail({_internal.ID})");
                _arguments.Add("contact", _ExistingContact);
                ApiQueueJob _findAccountById = new ApiQueueJob(_internal.ID,queue_FindAccountByID, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findAccountById);
            }
            else
            {
                //try searching by domain
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindContactByEmail({_internal.ID})");
                _arguments.Add("contact", null);
                ApiQueueJob _findAccountByDomain = new ApiQueueJob(_internal.ID, queue_FindAccountByDomain, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findAccountByDomain);
            }
        }
        private void queue_FindResourceByEmail(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Resource _ExistingResource = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindResourceByEmail({_internal.ID})");
            try
            {
                _ExistingResource = _atwsAPIClient.FindResourceByEmail(_internal.Message.Sender.EmailAddress.Address);
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindResourceByEmail({_internal.ID})");
                throw new ApiQueueException($"queue_FindResourceByEmail({_internal.ID})", _included);
            }
            if (_ExistingResource != null)
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindResourceByEmail({_internal.ID})");
            else
                //try searching by domain
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindResourceByEmail({_internal.ID})");
            //continue with creating a ticket note anyway.
            _arguments.Add("resource", _ExistingResource); // we really don't or shouldn't care of it the resource exists, this is just for impersonation.
            ApiQueueJob _createTicketNote = new ApiQueueJob(_internal.ID, queue_CreateTicketNote, _arguments, 5, new TimeSpan(0, 1, 0), true);
            _queueManager.Enqueue(_createTicketNote);
        }

        private void queue_FindAccountByID(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Contact _contact = (AutotaskPSA.Contact)_arguments["contact"];
            AutotaskPSA.Account _ExistingAccount = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindAccountByID({_internal.ID})");
            try
            {
                _ExistingAccount = _atwsAPIClient.FindAccountByID((int)_contact.AccountID);
            }
            catch(Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindAccountByID({_internal.ID})");
                throw new ApiQueueException($"queue_FindAccountByID({_internal.ID})", _included);
            }
            if (_ExistingAccount != null)
            {
                //we have an account stop looking and create a damn ticket
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindAccountByID({_internal.ID})");
                _arguments.Add("account", _ExistingAccount);
                ApiQueueJob _createTicket = new ApiQueueJob(_internal.ID, queue_CreateTicket, _arguments, 5, new TimeSpan(0, 1, 0), true);
                _queueManager.Enqueue(_createTicket);
            }
            else
            {
                //we have not found an account attached to sender WTF?! try by domain anyway?
                //try searching by domain
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindAccountByID({_internal.ID})");
                if (_arguments.ContainsKey("contact")) { _arguments.Remove("contact"); }
                ApiQueueJob _findAccountByDomain = new ApiQueueJob(_internal.ID, queue_FindAccountByDomain, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findAccountByDomain);
            }
        }
        private void queue_FindAccountByDomain(Dictionary<string,object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Account _ExistingAccount = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindAccountByDomainName({_internal.ID})");
            try
            {
                _ExistingAccount = _atwsAPIClient.FindAccountByDomain(_internal.Message.Sender.EmailAddress.Address.Split("@")[1]);
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindAccountByDomainName({_internal.ID})");
                throw new ApiQueueException($"queue_FindAccountByDomain({_internal.ID})", _included);
            }
            if (_ExistingAccount != null)
            {
                //we have an account stop looking and create a damn ticket
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindAccountByDomainName({_internal.ID})");
                _arguments.Add("account", _ExistingAccount);
                ApiQueueJob _createTicket = new ApiQueueJob(_internal.ID, queue_CreateTicket, _arguments, 5, new TimeSpan(0, 1, 0), true);
                _queueManager.Enqueue(_createTicket);
            }
            else
            {
                //we have not found an account attached to sender WTF?! RESORT TO DEFAULT ACCOUNT!!
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindAccountByDomainName({_internal.ID})");
                ApiQueueJob _findAccountByName = new ApiQueueJob(_internal.ID, queue_FindAccountByName, _arguments, 5, new TimeSpan(0, 1, 0), true);
                _queueManager.Enqueue(_findAccountByName);
            }
        }
        private void queue_FindAccountByName(Dictionary<string,object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Account _ExistingAccount = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_FindAccountByName({_internal.ID})");
            try
            {
                _ExistingAccount = _atwsAPIClient.FindAccountByName(_configuration.Autotask.Defaults.Ticket.Account); // when all else fails use Defaults!
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_FindAccountByName({_internal.ID})");
                throw new ApiQueueException($"queue_FindAccountByName({_internal.ID})", _included);
            }
            if (_ExistingAccount != null)
            {
                //we have an account stop lookng and create a damn ticket
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_FindAccountByName({_internal.ID})");
                _arguments.Add("account", _ExistingAccount);
                ApiQueueJob _createTicket = new ApiQueueJob(_internal.ID, queue_CreateTicket, _arguments, 5, new TimeSpan(0, 1, 0), true);
                _queueManager.Enqueue(_createTicket);
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_FindAccountByName({_internal.ID})");
                throw new ApiQueueException($"queue_FindAccountByName({_internal.ID})");
            }
        }
        private void queue_MoveMessageToFolder(Dictionary<string,object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            MailFolder _destination = (MailFolder)_arguments["folderDestination"];
            Message _MovedMessage = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_MoveMessageToFolder({_internal.ID})");
            try
            {
                _MovedMessage = _graphAPIClient.Users[_configuration.MailBox].MailFolders[_internal.Message.ParentFolderId].Messages[_internal.Message.Id].Move(_destination.Id).Request().PostAsync().GetAwaiter().GetResult();
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_MoveMessageToFolder({_internal.ID})");
                throw new ApiQueueException($"queue_MoveMessageToFolder({_internal.ID})", _included);
            }
            if (_MovedMessage != null)
            {
                //mark it read
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_MoveMessageToFolder({_internal.ID})");
                _internal.Message = _MovedMessage;
                _arguments.Remove("message");
                _arguments.Remove("folderDestination");
                _arguments.Add("message", _internal);
                _arguments.Add("isread", true);
                ApiQueueJob _markMessageAs = new ApiQueueJob(_internal.ID, queue_MarkMessageAs, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_markMessageAs);
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_MoveMessageToFolder({_internal.ID})");
                throw new ApiQueueException($"queue_MoveMessageToFolder({_internal.ID})");
            }
        }
        private void queue_MarkMessageAs(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            bool _isread = (bool)_arguments["isread"];
            Message _MarkedMessage = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_MarkMessageAs({_internal.ID})");
            try
            {
                _internal.Message.IsRead = _isread;
                _MarkedMessage = _graphAPIClient.Users[_configuration.MailBox].Messages[_internal.Message.Id].Request().Select("IsRead").UpdateAsync(new Message { IsRead = _isread }).GetAwaiter().GetResult();
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_MarkMessageAs({_internal.ID})");
                throw new ApiQueueException($"queue_MarkMessageAs({_internal.ID})", _included);
            }
            if (_MarkedMessage != null)
            {
                //we marked it OK
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_MarkMessageAs({_internal.ID})");
                _arguments.Remove("isread");
                ApiQueueJob _getMessageAttachments = new ApiQueueJob(_internal.ID, queue_GetMessageAttachments, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_getMessageAttachments);
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_MarkMessageAs({_internal.ID})");
                throw new ApiQueueException($"queue_MarkMessageAs({_internal.ID})");
            }
        }
        private void queue_CreateTicketNote(Dictionary<string,object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Ticket _ticket = (AutotaskPSA.Ticket)_arguments["ticket"];
            AutotaskPSA.Resource _resource = (AutotaskPSA.Resource)_arguments["resource"];
            TicketNote _TicketNote = null;
            bool _flaggedInternal = false;

            if (_arguments.ContainsKey("internal"))
                _flaggedInternal = (bool)_arguments["internal"];
            string _publish = _flaggedInternal == false ? _configuration.Autotask.Defaults.Attachment.Publish : "Internal Project Team";

            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_CreateTicketNote({_internal.ID})");
            try
            {
                _TicketNote = _atwsAPIClient.CreateTicketNote(_ticket, _resource, _internal.Message.Subject.Replace((string)_ticket.TicketNumber, " ").Trim(), _internal.Message.Body.Content, _configuration.Autotask.Defaults.TicketNote.Type, _publish);
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_CreateTicketNote({_internal.ID})");
                throw new ApiQueueException($"queue_CreateTicketNote({_internal.ID})", _included);
            }
            if (_TicketNote != null)
            {
                //ticket created queue message move now - attachments?
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_CreateTicketNote({_internal.ID})");
                _arguments.Add("folderDestination", _processed_folder);
                _arguments.Add("ticketnote", _TicketNote);
                if (_arguments.ContainsKey("ticket"))
                    _arguments.Remove("ticket");
                ApiQueueJob _moveMessageToFolder = new ApiQueueJob(_internal.ID, queue_MoveMessageToFolder, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_moveMessageToFolder);
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_CreateTicketNote({_internal.ID})");
                throw new ApiQueueException($"queue_CreateTicketNote({_internal.ID})");
            }
        }

        internal SKImage SkiaCheckAndResizeImage(SKImage _image)
        {
            try
            {
                if (_image != null)
                {
                    if (_image.EncodedData.Size >= 5900000) //maybe this max size for the image resize should be configurable
                    {
                        using (SKBitmap _bitmap = SKBitmap.FromImage(_image))
                        {
                            if (_bitmap != null)
                            {
                                using (SKBitmap _resized = new SKBitmap(new SKImageInfo((int)(_image.Width * .50), (int)(_image.Height * .50))))
                                {
                                    if (_resized != null)
                                    {
                                        _bitmap.ScalePixels(_resized, SKFilterQuality.High);
                                        using (SKImage _new = SKImage.FromBitmap(_resized))
                                        {
                                            if (_new != null)
                                            {
                                                using (SKData _output = _new.Encode(SKEncodedImageFormat.Jpeg, 50))
                                                {
                                                    if (_output != null)
                                                        return SKImage.FromEncodedData(_output);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception _ex)
            {
                _ex = _ex;
            }
            return _image;
        }
        internal Microsoft.Graph.FileAttachment ZipCheckAndCompressAttachment(Microsoft.Graph.FileAttachment _Attachment)
        {
            if (_Attachment == null | String.IsNullOrWhiteSpace(_Attachment.Name) || _Attachment.ContentBytes == null)
                return _Attachment;
            try
            {
                Microsoft.Graph.FileAttachment _compressed = new Microsoft.Graph.FileAttachment()
                {
                    AdditionalData = _Attachment.AdditionalData,
                    ODataType = _Attachment.ODataType
                };
                using (MemoryStream _memStream = new MemoryStream())
                {
                    using (ZipArchive _archive = new ZipArchive(_memStream, ZipArchiveMode.Create, true))
                    {
                        ZipArchiveEntry _zipEntry = _archive.CreateEntry(_Attachment.Name, CompressionLevel.Optimal);
                        using (Stream _zipEntryStream = _zipEntry.Open())
                        {
                            using (StreamWriter _zipStreamWriter = new StreamWriter(_zipEntryStream))
                            {
                                _zipStreamWriter.BaseStream.Write(_Attachment.ContentBytes, 0, _Attachment.ContentBytes.Length);
                                _zipStreamWriter.Flush();
                            }
                        }
                    }
                    _compressed.Name = $"{_Attachment.Name}.zip";
                    _compressed.ContentType = _mimeDatabase.GetMimeTypeInfoFromExtension("zip").MimeType;
                    _compressed.ContentBytes = _memStream.ToArray();
                    return _compressed;
                }
            }
            catch (Exception _ex)
            {
                _ex = _ex;
            }
            return _Attachment;
        }

        private void queue_CreateAttachment(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            Microsoft.Graph.FileAttachment _attachment = (Microsoft.Graph.FileAttachment)_arguments["attachment"];
            AutotaskPSA.Ticket _ticket = null;
            AutotaskPSA.TicketNote _ticketnote = null;
            AutotaskPSA.Resource _resource = null;
            AutotaskPSA.Contact _contact = null;
            bool _flaggedInternal = false;
            
            if (_arguments.ContainsKey("ticket"))
                _ticket = (AutotaskPSA.Ticket)_arguments["ticket"];
            if (_arguments.ContainsKey("ticketnote"))
                _ticketnote = (AutotaskPSA.TicketNote)_arguments["ticketnote"];
            if (_arguments.ContainsKey("resource"))
                _resource = (AutotaskPSA.Resource)_arguments["resource"];
            if (_arguments.ContainsKey("contact"))
                _contact = (AutotaskPSA.Contact)_arguments["contact"];
            if (_arguments.ContainsKey("internal"))
                _flaggedInternal = (bool)_arguments["internal"];

            string _publish = _flaggedInternal == false ? _configuration.Autotask.Defaults.Attachment.Publish : "Internal Users Only";
            AutotaskPSA.Attachment _Attachment = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_CreateAttachment({_internal.ID})");
            try
            {
                if (_ticket != null)
                    _Attachment = _atwsAPIClient.CreateTicketAttachment(_ticket, _resource, _attachment.Name, _attachment.ContentBytes, _publish);//change this to be configurable(All users/internal)
                else if(_ticketnote != null)
                   _Attachment = _atwsAPIClient.CreateTicketNoteAttachment(_ticketnote, _resource, _attachment.Name, _attachment.ContentBytes, _publish);//change this to be configurable(allusers/internal
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_CreateAttachment({_internal.ID})");
                throw new ApiQueueException($"queue_CreateAttachment({_internal.ID})", _included);
            }
            if (_Attachment != null)
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_CreateAttachment({_internal.ID})");
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_CreateAttachment({_internal.ID})");
                throw new ApiQueueException($"queue_CreateAttachment({_internal.ID})");
            }
        }
        private void queue_GetMessageAttachments(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Ticket _ticket = null;
            AutotaskPSA.TicketNote _ticketnote = null;
            AutotaskPSA.Resource _resource = null;
            AutotaskPSA.Contact _contact = null;
            if (_arguments.ContainsKey("ticket"))
                _ticket = (AutotaskPSA.Ticket)_arguments["ticket"];
            if (_arguments.ContainsKey("ticketnote"))
                _ticketnote = (AutotaskPSA.TicketNote)_arguments["ticketnote"];
            if (_arguments.ContainsKey("resource"))
                _resource = (AutotaskPSA.Resource)_arguments["resource"]; 
            if (_arguments.ContainsKey("contact"))
                _contact = (AutotaskPSA.Contact)_arguments["contact"];
            IMessageAttachmentsCollectionPage _Attachments = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_GetMessageAttachments({_internal.ID})");
            try
            {
                IMessageRequestBuilder _msgRequest = _graphAPIClient.Users[_configuration.MailBox].Messages[_internal.Message.Id];
                _Attachments = _msgRequest.Attachments.Request().GetAsync().GetAwaiter().GetResult();
            }
            catch (Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_GetMessageAttachments({_internal.ID})");
                throw new ApiQueueException($"queue_GetMessageAttachments({_internal.ID})", _included);
            }
            while (_Attachments != null)
            {
                foreach (Microsoft.Graph.FileAttachment _Attachment in _Attachments)
                {
                    try
                    {
                        string _extension = System.IO.Path.GetExtension(_Attachment.Name).Replace(".", "");
                        MimeTypeDatabase.MimeTypeInfo _mimeInfo = _mimeDatabase.GetMimeTypeInfoFromExtension(_extension);
                        if (_mimeInfo != null)
                        {
                            if (_mimeInfo.Allowable == false)
                            {
                                _logger.LogInformation($"[{_configuration.MailBox}] - SKIPBADITEM - queue_GetMessageAttachments({_internal.ID})");
                                continue;
                            }
                        }
                        using (SkiaSharp.SKImage _image = SkiaSharp.SKImage.FromEncodedData(_Attachment.ContentBytes))
                        {
                            if (_image != null)
                            {
                                if (_image.Width >= _configuration.Autotask.Defaults.Attachment.Image.MinWidth && _image.Height >= _configuration.Autotask.Defaults.Attachment.Image.MinHeight)
                                {
                                    if (_configuration.Autotask.Defaults.Attachment.ResizeLargeImages)
                                    {
                                        using (SKImage _checked = SkiaCheckAndResizeImage(_image))
                                        {
                                            _Attachment.ContentBytes = _checked.EncodedData.ToArray();
                                        }
                                    }
                                }
                                else
                                {
                                    _logger.LogInformation($"[{_configuration.MailBox}] - SKIPSMALLIMG - queue_GetMessageAttachments({_internal.ID})");
                                    continue;
                                }
                            }
                        }
                        if (_Attachment.ContentBytes.Length >= 5900000 && _configuration.Autotask.Defaults.Attachment.CompressLargeItems)
                        {
                            Microsoft.Graph.FileAttachment _compressed = ZipCheckAndCompressAttachment(_Attachment);
                            if (_compressed.ContentBytes.Length < _Attachment.ContentBytes.Length)
                            {
                                _Attachment.Name = _compressed.Name;
                                _Attachment.ContentType = _compressed.ContentType;
                                _Attachment.ContentBytes = _compressed.ContentBytes;
                            }
                        }
                        if (_Attachment.ContentBytes.Length <= 5900000)
                        {
                            //we got it small enough lets Queue Uploading
                            if (_arguments.ContainsKey("attachment"))
                                _arguments.Remove("attachment");
                            _arguments.Add("attachment", _Attachment);
                            ApiQueueJob _createAttachment = new ApiQueueJob(_internal.ID, queue_CreateAttachment, _arguments, 3, new TimeSpan(0, 5, 0));
                            _queueManager.Enqueue(_createAttachment);
                        }
                        else
                            _logger.LogInformation($"[{_configuration.MailBox}] - LARGEITEM - queue_GetMessageAttachments({_internal.ID})");
                    }
                    catch (Exception _ex)
                    {
                        Exception _included = _ex;
                        if (_ex.InnerException != null)
                            _included = _ex.InnerException;
                        _logger.LogError(_included,$"[{_configuration.MailBox}] - ERRORITEM - queue_GetMessageAttachments({_internal.ID})");
                    }
                }
                if (_Attachments.NextPageRequest != null)
                    _Attachments = _Attachments.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                else
                    _Attachments = null;
            }
            _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_GetMessageAttachments({_internal.ID})");
        }
        private void queue_CreateTicket(Dictionary<string, object> _arguments)
        {
            InternalMessage _internal = (InternalMessage)_arguments["message"];
            AutotaskPSA.Account _account = (AutotaskPSA.Account)_arguments["account"];
            AutotaskPSA.Contact _contact = (AutotaskPSA.Contact)_arguments["contact"];
            Ticket _Ticket = null;
            _logger.LogInformation($"[{_configuration.MailBox}] - ENTRY - queue_CreateTicket({_internal.ID})");
            try
            {
                _Ticket = _atwsAPIClient.CreateTicket(_account, _contact, _internal.Message.Subject, _internal.Message.Body.Content, (DateTime.Now + _configuration.Autotask.Defaults.Ticket.DueDateOffset), _configuration.Autotask.Defaults.Ticket.Source, _configuration.Autotask.Defaults.Ticket.Status, _configuration.Autotask.Defaults.Ticket.Priority, _configuration.Autotask.Defaults.Ticket.Queue, _configuration.Autotask.Defaults.Ticket.WorkType);
            }
            catch(Exception _ex)
            {
                Exception _included = _ex;
                if (_ex.InnerException != null)
                    _included = _ex.InnerException;
                _logger.LogInformation($"[{_configuration.MailBox}] - EXCEPTION - queue_CreateTicket({_internal.ID})");
                throw new ApiQueueException($"queue_CreateTicket({_internal.ID})", _included);
            }
            if (_Ticket != null && !string.IsNullOrWhiteSpace((string)_Ticket.TicketNumber))
            {
                //ticket created queue message move now - attachments?
                _logger.LogInformation($"[{_configuration.MailBox}] - SUCCESS - queue_CreateTicket({_internal.ID})");
                _arguments.Add("folderDestination", _processed_folder);
                _arguments.Add("ticket", _Ticket);
                ApiQueueJob _moveMessageToFolder = new ApiQueueJob(_internal.ID, queue_MoveMessageToFolder, _arguments, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_moveMessageToFolder);
            }
            else
            {
                _logger.LogInformation($"[{_configuration.MailBox}] - FAIL - queue_CreateTicket({_internal.ID})");
                throw new ApiQueueException($"queue_CreateTicket({_internal.ID})");
            }
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        private async System.Threading.Tasks.Task OnMessageReceivedAsync(Message message)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            InternalMessage _internal = new InternalMessage() { ID = Sha256Sum(message.Id), Message = message };
            if (!_queueManager.HasJobWithMessageID(_internal.ID))
            {
                Dictionary<string, object> _findTicketJobArgs = new Dictionary<string, object>();
                _logger.LogInformation($"[{_configuration.MailBox}] - OnMessageReceivedAsync({_internal.ID})");
                string _messageTag = ExtractSubjectTag(_internal.Message.Subject);
                bool _messageTagIsInternalFunction = false;
                if (!String.IsNullOrWhiteSpace(_messageTag))
                {
                    switch(_messageTag.ToLower())
                    {
                        case "internal":
                            _findTicketJobArgs.Add("internal", true);
                            _messageTagIsInternalFunction = true;
                            break;
                    }
                }

                if (_messageTagIsInternalFunction)
                    _internal.Message.Subject = _internal.Message.Subject.Replace($"[{_messageTag}]", "").Trim();
                _findTicketJobArgs.Add("message", _internal);
                ApiQueueJob _findTicketJob = new ApiQueueJob(_internal.ID, queue_FindTicketByNumber, _findTicketJobArgs, 5, new TimeSpan(0, 1, 0));
                _queueManager.Enqueue(_findTicketJob);
            }
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
        protected internal string ExtractSubjectTag(string data)
        {
            const string _ACS_TagRegEx = @"(\[[a-zA-Z]*\])";
            string _searchResult = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(data))
                {
                    MatchCollection _matches = Regex.Matches(data, _ACS_TagRegEx, RegexOptions.Multiline);
                    if (_matches.Count > 0)
                    {
                        _searchResult = _matches.First().Value.Replace("[", "").Replace("]", "");
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
