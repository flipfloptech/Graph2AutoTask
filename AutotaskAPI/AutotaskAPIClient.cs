using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;

namespace AutotaskPSA
{
    public class AutotaskAPIClient
    {
        private readonly ATWSSoapChannel _atwsServicesClient = null;
        private readonly BasicHttpBinding _atwsHttpBinding = null;
        private readonly ChannelFactory<ATWSSoapChannel> _atwsChannelFactory = null;
        private readonly string _atwsUsername = string.Empty;
        private readonly string _atwsPassword = string.Empty;
        private readonly string _atwsIntegrationKey = string.Empty;
        private readonly AutotaskIntegrations _atwsIntegrations = new AutotaskIntegrations();
        private readonly ATWSZoneInfo _atwsZoneInfo = null;
        private GetFieldInfoResponse account_fieldInfo = null;
        private getUDFInfoResponse account_udfInfo = null;
        private GetFieldInfoResponse ticket_fieldInfo = null;
        private GetFieldInfoResponse ticketNote_fieldInfo = null;
        private GetFieldInfoResponse attachmentInfo_fieldInfo = null;
        private List<AllocationCode> allocationCodes_WorkType = new List<AllocationCode>();
        private const string _AT_TicketRegEx = "^[T][0-9]{8}[.][0-9]{4}$";
        public AutotaskAPIClient(string Username, string Password, AutotaskIntegrationKey IntegrationKey)
        {
            if (string.IsNullOrWhiteSpace(Username))
            {
                throw new ArgumentNullException("Username cannot by empty");
            }

            if (string.IsNullOrWhiteSpace(Password))
            {
                throw new ArgumentNullException("Password cannot by empty");
            }

            string _integrationKey = IntegrationKey.GetIntegrationKey();
            if (string.IsNullOrWhiteSpace(_integrationKey))
            {
                throw new ArgumentNullException("IntegrationKey cannot by empty");
            }

            _atwsUsername = Username;
            _atwsPassword = Password;
            _atwsIntegrationKey = _integrationKey;
            _atwsIntegrations.IntegrationCode = _integrationKey;
            NetworkCredential _credential = new NetworkCredential(_atwsUsername, _atwsPassword);
            try
            {
                _atwsHttpBinding = new BasicHttpBinding(BasicHttpSecurityMode.Transport)
                {
                    MaxBufferSize = int.MaxValue,
                    MaxReceivedMessageSize = int.MaxValue
                };
                _atwsHttpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Basic;
                _atwsChannelFactory = new ChannelFactory<ATWSSoapChannel>(_atwsHttpBinding, new EndpointAddress(new Uri("https://webservices.autotask.net/ATServices/1.6/atws.asmx")));
                _atwsChannelFactory.Credentials.UserName.UserName = _atwsUsername;
                _atwsChannelFactory.Credentials.UserName.Password = _atwsPassword;
                _atwsServicesClient = _atwsChannelFactory.CreateChannel();
                _atwsServicesClient.Open();
                _atwsZoneInfo = _atwsServicesClient.getZoneInfo(_atwsUsername);
                if (_atwsZoneInfo.ErrorCode >= 0)
                {
                    _atwsServicesClient.Close();
                    _atwsChannelFactory.Close();
                    _atwsChannelFactory = new ChannelFactory<ATWSSoapChannel>(_atwsHttpBinding, new EndpointAddress(new Uri(_atwsZoneInfo.URL)));
                    _atwsChannelFactory.Credentials.UserName.UserName = _atwsUsername;
                    _atwsChannelFactory.Credentials.UserName.Password = _atwsPassword;
                    _atwsServicesClient = _atwsChannelFactory.CreateChannel();
                    _atwsServicesClient.Open();
                    if (!Preload())
                    {
                        throw new Exception("Error with data Preload()");
                    }
                }
                else
                {
                    throw new Exception("Error with getZoneInfo()");
                }
            }
            catch
            {
                throw new ArgumentException("Username must be invalid, unable to get ZoneInfo");
            }
        }

        ~AutotaskAPIClient()
        {
            try
            {
                if (_atwsServicesClient != null && _atwsServicesClient.State != CommunicationState.Closed)
                {
                    _atwsServicesClient.Close();
                }

                if (_atwsChannelFactory != null && _atwsChannelFactory.State != CommunicationState.Closed)
                {
                    _atwsChannelFactory.Close();
                }
            }
            catch
            { }
        }

        public bool Preload()
        {
            try
            {
                ticket_fieldInfo = _atwsServicesClient.GetFieldInfo(new GetFieldInfoRequest(_atwsIntegrations, "Ticket"));
                ticketNote_fieldInfo = _atwsServicesClient.GetFieldInfo(new GetFieldInfoRequest(_atwsIntegrations, "TicketNote"));
                attachmentInfo_fieldInfo = _atwsServicesClient.GetFieldInfo(new GetFieldInfoRequest(_atwsIntegrations, "AttachmentInfo"));
                account_fieldInfo = _atwsServicesClient.GetFieldInfo(new GetFieldInfoRequest(_atwsIntegrations, "Account"));
                account_udfInfo = _atwsServicesClient.getUDFInfo(new getUDFInfoRequest(_atwsIntegrations,"Account"));
                allocationCodes_WorkType = UpdateUseTypeOneAllocationCodes();
            }
            catch(Exception _ex)
            {
                _ex = _ex;
                return false;
            }
            return true;
        }

        public Attachment CreateTicketNoteAttachment(TicketNote ParentTicketNote, Resource ImpersonateResource, string Title, byte[] Data, string Publish)
        {
            Attachment _result = null;
            if (string.IsNullOrWhiteSpace(Title))
                Title = "No Name";
            if (ParentTicketNote == null || (Data == null || Data.Length <= 0 || Data.Length > 6000000) || string.IsNullOrWhiteSpace(Publish))
            {
                throw new ArgumentNullException();
            }

            try
            {
                _result = new Attachment()
                {
                    Info = new AttachmentInfo()
                    {
                        Type = "FILE_ATTACHMENT",
                        ParentType = 23,
                        ParentID = ParentTicketNote.id,
                        FullPath = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length),
                        Title = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length),
                        Publish = Convert.ToInt32(PickListValueFromField(attachmentInfo_fieldInfo.GetFieldInfoResult, "Publish", Publish)),
                    },
                    Data = Data
                };
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new ArgumentException("Unable to build attachment request. Please review InnerException.", _ex);
            }
            try
            {
                if (_result != null)
                {
                    _atwsIntegrations.ImpersonateAsResourceID = ImpersonateResource != null ? (int)ImpersonateResource.id : 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = _atwsIntegrations.ImpersonateAsResourceID == 0 ? false : true;
                    CreateAttachmentResponse _createResponse = _atwsServicesClient.CreateAttachment(new CreateAttachmentRequest(_atwsIntegrations, _result));
                    _atwsIntegrations.ImpersonateAsResourceID = 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = false;
                    if (_createResponse.CreateAttachmentResult > 0)
                    {
                        GetAttachmentResponse _getResponse = _atwsServicesClient.GetAttachment(new GetAttachmentRequest(_atwsIntegrations, _createResponse.CreateAttachmentResult));
                        if (_getResponse.GetAttachmentResult != null)
                            _result = _getResponse.GetAttachmentResult;
                        else
                            throw new CommunicationException("AutotaskAPIClient.CreateTickeNoteAttachment() UNKNOWN ERROR");
                    }
                }
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new CommunicationException("Unable to create ticket note attachment. Please review InnerException.", _ex);
            }
            return _result;
        }

        public Attachment CreateTicketAttachment(Ticket ParentTicket, Resource ImpersonateResource, string Title, byte[] Data, string Publish)
        {
            Attachment _result = null;
            if (string.IsNullOrWhiteSpace(Title))
                Title = "No Name";
            if (ParentTicket == null || (Data == null || Data.Length <= 0 || Data.Length > 6000000) || string.IsNullOrWhiteSpace(Publish))
            {
                throw new ArgumentNullException();
            }

            try
            {
                _result = new Attachment()
                {
                    //ParentTypes = /Account(1), Ticket(3), or Project(4).
                    Info = new AttachmentInfo()
                    {
                        Type = "FILE_ATTACHMENT",
                        ParentType = 4,
                        ParentID = ParentTicket.id,
                        FullPath = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length),
                        Title = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length),
                        Publish = Convert.ToInt32(PickListValueFromField(attachmentInfo_fieldInfo.GetFieldInfoResult, "Publish", Publish)),
                    },
                    Data = Data
                };
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new ArgumentException("Unable to build ticket attachment request. Please review InnerException.", _ex);
            }
            try
            {
                if (_result != null)
                {
                    _atwsIntegrations.ImpersonateAsResourceID = ImpersonateResource != null ? (int)ImpersonateResource.id : 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = _atwsIntegrations.ImpersonateAsResourceID == 0 ? false : true;
                    CreateAttachmentResponse _createResponse = _atwsServicesClient.CreateAttachment(new CreateAttachmentRequest(_atwsIntegrations, _result));
                    _atwsIntegrations.ImpersonateAsResourceID = 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = false;
                    if (_createResponse.CreateAttachmentResult > 0)
                    { 
                        GetAttachmentResponse _getResponse = _atwsServicesClient.GetAttachment(new GetAttachmentRequest(_atwsIntegrations, _createResponse.CreateAttachmentResult));
                        if (_getResponse.GetAttachmentResult != null)
                            _result = _getResponse.GetAttachmentResult;
                        else
                            throw new CommunicationException("AutotaskAPIClient.CreateTickeAttachment() UNKNOWN ERROR");
                    }
                }
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new CommunicationException("Unable to create ticket attachment. Please review InnerException.", _ex);
            }
            return _result;
        }
        public TicketNote CreateTicketNote(Ticket Ticket, Resource ImpersonateResource, string Title, string Description, string Type, string Publish)
        {
            TicketNote _result = null;
            if (String.IsNullOrWhiteSpace(Title))
                Title = "Empty subject line";
            if (string.IsNullOrWhiteSpace(Description))
                Description = "Empty message contents";
            if (Ticket == null || string.IsNullOrWhiteSpace(Type) || string.IsNullOrWhiteSpace(Publish))
            {
                throw new ArgumentNullException();
            }
            try
            {
                _result = new TicketNote()
                {
                    TicketID = Convert.ToInt32(Ticket.id),
                    Title = Title.Substring(0, Title.Length > 250 ? 250 : Title.Length),
                    Description = Description.Substring(0, Description.Length > 3200 ? 3200 : Description.Length),
                    NoteType = Convert.ToInt32(PickListValueFromField(ticketNote_fieldInfo.GetFieldInfoResult, "NoteType", Type)),
                    Publish = Convert.ToInt32(PickListValueFromField(ticketNote_fieldInfo.GetFieldInfoResult, "Publish", Publish)),
                };
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new ArgumentException("Unable to build ticketnote request. Please review InnerException.", _ex);
            }
            try
            {
                if (_result != null)
                {
                    _atwsIntegrations.ImpersonateAsResourceID = ImpersonateResource != null ? (int)ImpersonateResource.id : 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = _atwsIntegrations.ImpersonateAsResourceID == 0 ? false : true;
                    createResponse _response = _atwsServicesClient.create(new createRequest(_atwsIntegrations, new Entity[] { _result }));
                    _atwsIntegrations.ImpersonateAsResourceID = 0;
                    _atwsIntegrations.ImpersonateAsResourceIDSpecified = false;
                    if (_response.createResult.ReturnCode > 0 && _response.createResult.EntityResults.Length > 0)
                        _result = (TicketNote)_response.createResult.EntityResults[0];
                    else if(_response.createResult.ReturnCode <= 0 && _response.createResult.EntityResults.Length == 0 && _response.createResult.Errors.Count() > 0)
                        throw new CommunicationException(_response.createResult.Errors[0].Message);
                    else
                        throw new CommunicationException("AutotaskAPIClient.CreateTicketNote() UNKNOWN ERROR");
                }
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new CommunicationException("Unable to create ticketnote. Please review InnerException.", _ex);
            }
            return _result;
        }
        public Ticket UpdateTicket(Ticket OriginalTicket, string Title = null, string Description = null, DateTimeOffset? DueDate = null, string Source = null, string Status = null, string Priority = null, string Queue = null, string WorkType = null)
        {
            Ticket _result = null;
            Ticket _UpdatedTicket = null;
            if (OriginalTicket == null)
                throw new ArgumentNullException("OriginalTicket missing");
            _UpdatedTicket = OriginalTicket;
            try
            {
                if (String.IsNullOrWhiteSpace(Title) == false)
                    _UpdatedTicket.Title = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length);

                if (String.IsNullOrWhiteSpace(Description) == false)
                    _UpdatedTicket.Description = Description.Substring(0, Description.Length > 8000 ? 8000 : Description.Length);
                
                if (DueDate != null)
                    _UpdatedTicket.DueDateTime = DueDate.Value.LocalDateTime;
                
                if (String.IsNullOrWhiteSpace(Source) == false)
                    _UpdatedTicket.Source = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Source", Source));
                
                if (String.IsNullOrWhiteSpace(Status) == false)
                    _UpdatedTicket.Status = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Status", Status));

                if (String.IsNullOrWhiteSpace(Priority) == false)
                    _UpdatedTicket.Priority = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Priority", Priority));

                if (String.IsNullOrWhiteSpace(Queue) == false)
                    _UpdatedTicket.QueueID = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "QueueID", Queue));

                if (String.IsNullOrWhiteSpace(WorkType) == false)
                    _UpdatedTicket.AllocationCodeID = allocationCodes_WorkType.Exists(_code => Convert.ToString(_code.Name) == WorkType) ? allocationCodes_WorkType.First(_code => Convert.ToString(_code.Name) == WorkType).id : new long?();
            }
            catch (Exception _ex)
            {
                _UpdatedTicket = null;
                throw new ArgumentException("Unable to update ticket request. Please review InnerException.", _ex);
            }
            try
            {
                if (_UpdatedTicket != null)
                {
                    updateResponse _response = _atwsServicesClient.update(new updateRequest(_atwsIntegrations, new Entity[] { _UpdatedTicket }));
                    if (_response.updateResult.ReturnCode > 0 && _response.updateResult.EntityResults.Length > 0)
                        _result = (Ticket)_response.updateResult.EntityResults[0];
                    else if (_response.updateResult.ReturnCode <= 0 && _response.updateResult.EntityResults.Length == 0 && _response.updateResult.Errors.Count() > 0)
                        throw new CommunicationException(_response.updateResult.Errors[0].Message);
                    else
                        throw new CommunicationException("AutotaskAPIClient.UpdateTicket() UNKNOWN ERROR");
                }
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new CommunicationException("Unable to update ticket. Please review InnerException.", _ex);
            }
            return _result;
        }

        public Ticket CreateTicket(Account Account, Contact Contact, string Title, string Description, DateTimeOffset DueDate, string Source, string Status, string Priority, string Queue, string WorkType)
        {
            Ticket _result = null;
            if (String.IsNullOrWhiteSpace(Title))
                Title = "Empty subject line";
            if (string.IsNullOrWhiteSpace(Description))
                Description = "Empty message contents";
            if (Account == null || DueDate == null || string.IsNullOrWhiteSpace(Source) || string.IsNullOrWhiteSpace(Status) || string.IsNullOrWhiteSpace(Priority) || string.IsNullOrWhiteSpace(Queue) || string.IsNullOrWhiteSpace(WorkType))
            {
                throw new ArgumentNullException();
            }

            try
            {
                _result = new Ticket()
                {
                    AccountID = Account.id,
                    DueDateTime = DueDate.LocalDateTime,
                    Priority = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Priority", Priority)),
                    Status = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Status", Status)),
                    QueueID = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "QueueID", Queue)),
                    Source = Convert.ToInt32(PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "Source", Source)),
                    AllocationCodeID = allocationCodes_WorkType.Exists(_code => Convert.ToString(_code.Name) == WorkType) ? allocationCodes_WorkType.First(_code => Convert.ToString(_code.Name) == WorkType).id : new long?(),
                    ContactID = (Contact != null) ? Convert.ToInt32(Contact.id) : new int?(), // ternary null bug work around: new int?()
                    Title = Title.Substring(0, Title.Length > 255 ? 255 : Title.Length),
                    Description = Description.Substring(0, Description.Length > 8000 ? 8000 : Description.Length),
                };
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new ArgumentException("Unable to build ticket request. Please review InnerException.", _ex);
            }
            try
            {
                if (_result != null)
                {
                    createResponse _response = _atwsServicesClient.create(new createRequest(_atwsIntegrations, new Entity[] { _result }));
                    if (_response.createResult.ReturnCode > 0 && _response.createResult.EntityResults.Length > 0)
                        _result = (Ticket)_response.createResult.EntityResults[0];
                    else if (_response.createResult.ReturnCode <= 0 && _response.createResult.EntityResults.Length == 0 && _response.createResult.Errors.Count() > 0)
                        throw new CommunicationException(_response.createResult.Errors[0].Message);
                    else
                        throw new CommunicationException("AutotaskAPIClient.CreateTicket() UNKNOWN ERROR");
                }
            }
            catch (Exception _ex)
            {
                _result = null;
                throw new CommunicationException("Unable to create ticket. Please review InnerException.", _ex);
            }
            return _result;
        }

        public Ticket[] FindTicketByDateRange(DateTimeOffset DateStart, DateTimeOffset DateEnd)
        {
            Ticket[] _results = null;
            List<Ticket> _resultBuilder = new List<Ticket>();
            if (DateStart == null && DateEnd == null)
            {
                throw new ArgumentNullException();
            }

            if (DateStart >= DateEnd)
            {
                throw new ArgumentOutOfRangeException();
            }

            StringBuilder strQuery = new StringBuilder();
            StringBuilder strQueryStart = new StringBuilder();
            StringBuilder strQueryEnd = new StringBuilder();
            List<StringBuilder> strConditions = new List<StringBuilder>();
            strQueryStart.Append("<queryxml version=\"1.0\">");
            strQueryStart.Append("<entity>Ticket</entity>");
            strQueryStart.Append("<query>");

            strConditions.Add(new StringBuilder());
            strConditions.Last().Append("<condition><field>CreateDate<expression op=\"GreaterThanorEquals\">");
            strConditions.Last().Append(DateStart);
            strConditions.Last().Append("</expression></field></condition>");

            strConditions.Add(new StringBuilder());
            strConditions.Last().Append("<condition><field>CreateDate<expression op=\"LessThanorEquals\">");
            strConditions.Last().Append(DateEnd);
            strConditions.Last().Append("</expression></field></condition>");

            strQueryEnd.Append("</query></queryxml>");

            strQuery.Append(strQueryStart);
            foreach (StringBuilder _condition in strConditions)
            {
                strQuery.Append(_condition);
            }
            strQuery.Append(strQueryEnd);

            queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strQuery.ToString()));
            while (respResource.queryResult.ReturnCode > 0)
            {
                Ticket[] _temp = new Ticket[respResource.queryResult.EntityResults.Count()];
                Array.Copy(respResource.queryResult.EntityResults, 0, _temp, 0, respResource.queryResult.EntityResults.Count());
                _resultBuilder.AddRange(_temp);
                if (respResource.queryResult.EntityResults.Length == 500)
                {// try for more
                    if (strConditions.Count == 3)
                    {
                        strConditions.Remove(strConditions.Last());
                    }

                    strConditions.Add(new StringBuilder());
                    strConditions.Last().Append("<condition><field>id<expression op=\"GreaterThan\">");
                    strConditions.Last().Append(respResource.queryResult.EntityResults.Last().id);
                    strConditions.Last().Append("</expression></field></condition>");
                    strQuery.Clear();
                    strQuery.Append(strQueryStart);
                    foreach (StringBuilder _condition in strConditions)
                    {
                        strQuery.Append(_condition);
                    }

                    strQuery.Append(strQueryEnd);
                    respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strQuery.ToString()));
                }
                else
                {
                    break; // we got em all but returncode will never be > 0 unless we set it so just break out.
                }
            }
            _results = _resultBuilder.ToArray();
            return _results;

        }

        public Account UpdateAccountPickListUDF(Account account, string FieldName, string FieldValue)
        {
            try {
                foreach(var _field in account.UserDefinedFields)
                {
                    if (_field.Name == FieldName)
                    {
                        _field.Value = PickListValueFromField(account_udfInfo.getUDFInfoResult,FieldName,FieldValue);
                        break;
                    }
                }
                updateResponse _updateResp = _atwsServicesClient.update(new updateRequest(_atwsIntegrations, new Entity[1] {account}));
                if (_updateResp.updateResult.ReturnCode == 1)
                    return (Account)_updateResp.updateResult.EntityResults[0];  
                else
                    return null;
            }
            catch(Exception _ex)
            {
                return null;
            }
        }
        public Account[] FindAccountsByFieldValue(string FieldName, string FieldOperator = "contains", string FieldValue = null, bool isUDF = false)
        {
            Dictionary<int, decimal> _counts = new Dictionary<int, decimal>();
            List<Account>_result = new List<Account>();
            // Query Resource to get the owner of the lead
            StringBuilder strQuery = new StringBuilder();
            StringBuilder strQueryStart = new StringBuilder();
            StringBuilder strQueryEnd = new StringBuilder();
            List<StringBuilder> strConditions = new List<StringBuilder>();
            try
            {
                strQueryStart.Append("<queryxml version=\"1.0\">");
                strQueryStart.Append("<entity>Account</entity>");
                strQueryStart.Append("<query>");

                strConditions.Add(new StringBuilder());
                strConditions.Last().Append($"<condition><field udf=\"{isUDF.ToString().ToLower()}\">{FieldName}<expression op=\"{FieldOperator}\">");
                strConditions.Last().Append(FieldValue);
                strConditions.Last().Append("</expression></field></condition>");

                strQueryEnd.Append("</query></queryxml>");

                strQuery.Append(strQueryStart);
                foreach (StringBuilder _condition in strConditions)
                {
                    strQuery.Append(_condition);
                }
                strQuery.Append(strQueryEnd);
                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strQuery.ToString()));

                while (respResource.queryResult.ReturnCode > 0)
                {
                    Account[] _temp = new Account[respResource.queryResult.EntityResults.Count()];
                    Array.Copy(respResource.queryResult.EntityResults, 0, _temp, 0, respResource.queryResult.EntityResults.Count());
                    _result.AddRange(_temp);
                    if (respResource.queryResult.EntityResults.Length == 500)
                    {// try for more
                        if (strConditions.Count == 2)
                        {
                            strConditions.Remove(strConditions.Last());
                        }

                        strConditions.Add(new StringBuilder());
                        strConditions.Last().Append("<condition><field>id<expression op=\"GreaterThan\">");
                        strConditions.Last().Append(respResource.queryResult.EntityResults.Last().id);
                        strConditions.Last().Append("</expression></field></condition>");
                        strQuery.Clear();
                        strQuery.Append(strQueryStart);
                        foreach (StringBuilder _condition in strConditions)
                        {
                            strQuery.Append(_condition);
                        }
                        strQuery.Append(strQueryEnd);
                        respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strQuery.ToString()));
                    }
                    else
                    {
                        break; // we got em all but returncode will never be > 0 unless we set it so just break out.
                    }
                }
                return _result.ToArray();
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindAccountsByFieldValue", _ex);
            }
        }
        public string LookupTicketStatusByValue(int value)
        {
            string _result = null;
            try
            {
                _result = PickListLabelFromValue(ticket_fieldInfo.GetFieldInfoResult, "Status", Convert.ToString(value));
            }
            catch (Exception)
            {
                _result = null;
            }
            return _result;
        }
        public string LookupTicketQueueByValue(int value)
        {
            string _result = null;
            try
            {
                _result = PickListLabelFromValue(ticket_fieldInfo.GetFieldInfoResult, "QueueID", Convert.ToString(value));
            }
            catch (Exception)
            {
                _result = null;
            }
            return _result;
        }
        public string LookupTicketQueueValueByName(string Name)
        {
            string _result = null;
            try
            {
                _result = PickListValueFromField(ticket_fieldInfo.GetFieldInfoResult, "QueueID", Name);
            }
            catch (Exception)
            {
                _result = null;
            }
            return _result;
        }
        public string LookupAccountUDFByValue(string udfName, int value)
        {
            string _result = null;
            try
            {
                
                _result = PickListLabelFromValue(account_udfInfo.getUDFInfoResult,udfName,Convert.ToString(value));
            }
            catch (Exception)
            {
                _result = null;
            }
            return _result;
        }
        public string LookupTicketPriorityByValue(int value)
        {
            string _result = null;
            try
            {
                _result = PickListLabelFromValue(ticket_fieldInfo.GetFieldInfoResult, "Priority", Convert.ToString(value));
            }
            catch (Exception)
            {
                _result = null;
            }
            return _result;
        }

        private List<AllocationCode> UpdateUseTypeOneAllocationCodes()
        {
            List<AllocationCode> _result = null;

            StringBuilder strResource = new StringBuilder();
            strResource.Append("<queryxml version=\"1.0\">");
            strResource.Append("<entity>AllocationCode</entity>");
            strResource.Append("<query>");
            strResource.Append("<condition><field>UseType<expression op=\"equals\">");
            strResource.Append(Convert.ToString(1));
            strResource.Append("</expression></field></condition>");
            strResource.Append("</query></queryxml>");

            queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

            if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
            {
                _result = new List<AllocationCode>(Array.ConvertAll(respResource.queryResult.EntityResults, new Converter<Entity, AllocationCode>(EntityToAllocationCode)));
            }
            return _result;
        }
        private AllocationCode EntityToAllocationCode(Entity _entity)
        {
            return (AllocationCode)_entity;
        }
        public Ticket FindTicketByNumber(string ticketNumber)
        {
            Ticket _result = null;
            StringBuilder strResource = new StringBuilder();
            if (string.IsNullOrWhiteSpace(ticketNumber))
            {
                throw new ArgumentNullException();
            }

            if (!IsValidAutoTaskTicket(ticketNumber))
            {
                throw new ArgumentException();
            }
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Ticket</entity>");
                strResource.Append("<query>");
                strResource.Append("<field>TicketNumber<expression op=\"equals\">");
                strResource.Append(Convert.ToString(ticketNumber));
                strResource.Append("</expression></field>");
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    _result = (Ticket)respResource.queryResult.EntityResults[0];
                }
                return _result;
            }
            catch(Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindTicketByNumber",_ex);
            }
        }

        public Resource FindResource(string resourceUserName)
        {
            Resource _result = null;

            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            strResource.Append("<queryxml version=\"1.0\">");
            strResource.Append("<entity>Resource</entity>");
            strResource.Append("<query>");
            strResource.Append("<field>UserName<expression op=\"equals\">");
            strResource.Append(resourceUserName);
            strResource.Append("</expression></field>");
            strResource.Append("</query></queryxml>");

            queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

            if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
            {
                // Get the ID for the resource
                _result = (Resource)respResource.queryResult.EntityResults[0];
            }

            return _result;
        }
        public Resource FindResourceByEmail(string resourceEmail)
        {
            Resource _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            try { 
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Resource</entity>");
                strResource.Append("<query>");
                strResource.Append("<field>Email<expression op=\"equals\">");
                strResource.Append(resourceEmail);
                strResource.Append("</expression></field>");
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    // Get the ID for the resource
                    _result = (Resource)respResource.queryResult.EntityResults[0];
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindResourceByEmail", _ex);
            }
        }

        public Contact FindContactByEmail(string contactEmail)
        {
            Contact _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Contact</entity>");
                strResource.Append("<query>");
                strResource.Append("<condition><field>EMailAddress<expression op=\"equals\">");
                strResource.Append(contactEmail);
                strResource.Append("</expression></field></condition>");

                strResource.Append("<condition><field>Active<expression op=\"equals\">");
                strResource.Append("1");
                strResource.Append("</expression></field></condition>");
                
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    _result = (Contact)respResource.queryResult.EntityResults[0];
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindContactByEmail", _ex);
            }
        }
        public Contact FindContactByID(int ContactID)
        {
            Contact _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Contact</entity>");
                strResource.Append("<query>");
                strResource.Append("<condition><field>id<expression op=\"equals\">");
                strResource.Append(Convert.ToString(ContactID));
                strResource.Append("</expression></field></condition>");

                strResource.Append("<condition><field>Active<expression op=\"equals\">");
                strResource.Append("1");
                strResource.Append("</expression></field></condition>");
                
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    _result = (Contact)respResource.queryResult.EntityResults[0];
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindContactByID", _ex);
            }
        }
        public Contact FindFirstContactByDomain(string emailDomain)
        {
            Contact _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            strResource.Append("<queryxml version=\"1.0\">");
            strResource.Append("<entity>Contact</entity>");
            strResource.Append("<query>");
            strResource.Append("<field>EMailAddress<expression op=\"contains\">");
            strResource.Append(emailDomain);
            strResource.Append("</expression></field>");
            strResource.Append("</query></queryxml>");

            queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

            if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
            {
                _result = (Contact)respResource.queryResult.EntityResults[0];
            }

            return _result;
        }

        public Account FindAccountByID(int AccountID)
        {
            Account _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Account</entity>");
                strResource.Append("<query>");
                strResource.Append("<field>id<expression op=\"equals\">");
                strResource.Append(Convert.ToString(AccountID));
                strResource.Append("</expression></field>");
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    _result = (Account)respResource.queryResult.EntityResults[0];
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindAccountByID", _ex);
            }
        }

        public Account FindAccountByName(string AccountName)
        {
            Account _result = null;
            StringBuilder strResource = new StringBuilder();
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Account</entity>");
                strResource.Append("<query>");
                strResource.Append("<field>AccountName<expression op=\"equals\">");
                strResource.Append(AccountName);
                strResource.Append("</expression></field>");
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    _result = (Account)respResource.queryResult.EntityResults[0];
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindAccountByName", _ex);
            }
        }
        public Account FindAccountByDomain(string emailDomain)
        {
            Dictionary<int, decimal> _counts = new Dictionary<int, decimal>();
            Dictionary<int, decimal> _statistics = new Dictionary<int, decimal>();
            Account _result = null;
            // Query Resource to get the owner of the lead
            StringBuilder strResource = new StringBuilder();
            try
            {
                strResource.Append("<queryxml version=\"1.0\">");
                strResource.Append("<entity>Contact</entity>");
                strResource.Append("<query>");
                strResource.Append("<field>EMailAddress<expression op=\"contains\">");
                strResource.Append(emailDomain);
                strResource.Append("</expression></field>");
                strResource.Append("</query></queryxml>");

                queryResponse respResource = _atwsServicesClient.query(new queryRequest(_atwsIntegrations, strResource.ToString()));

                if (respResource.queryResult.ReturnCode > 0 && respResource.queryResult.EntityResults.Length > 0)
                {
                    int _total = respResource.queryResult.EntityResults.Count();

                    System.Threading.Tasks.Parallel.ForEach(respResource.queryResult.EntityResults, (_contact) =>
                    {
                        lock (_counts)
                        {
                            if (_counts.ContainsKey((int)((Contact)_contact).AccountID))
                            {
                                _counts[(int)((Contact)_contact).AccountID]++;
                            }
                            else
                            {
                                _counts.Add((int)((Contact)_contact).AccountID, 1);
                                _statistics.Add((int)((Contact)_contact).AccountID, 0);
                            }
                        }
                    });
                    System.Threading.Tasks.Parallel.ForEach(_counts, (_count) =>
                    {
                        _statistics[_count.Key] = Math.Round(((_count.Value / _total) * 100), 3);
                    });

                    int _AccountID = _statistics.OrderBy(key => key.Value).Reverse().First().Key;
                    _result = FindAccountByID(_AccountID);
                }

                return _result;
            }
            catch (Exception _ex)
            {
                throw new Exception("AutotaskAPIClient.FindAccountByDomain", _ex);
            }
        }

        protected internal static bool IsValidAutoTaskTicket(string ticketNumber)
        {
            bool _searchResult = false;
            if (!string.IsNullOrWhiteSpace(ticketNumber))
            {
                Match _match = Regex.Match(ticketNumber, _AT_TicketRegEx, RegexOptions.Singleline);
                if (_match.Success)
                {
                    _searchResult = true;
                }
            }
            return _searchResult;
        }
        protected internal static bool MSIsValidEmail(string email)
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



        /// <summary>
        /// Returns the value of a picklistitem given an array of fields and the field name that contains the picklist
        /// </summary>
        /// <param name="fields">array of Fields that contains the field to find</param>
        /// <param name="strField">name of the Field to search in</param>
        /// <param name="strPickListName">name of the picklist value to return the value from</param>
        /// <returns>value of the picklist item to search for</returns>
        protected static string PickListValueFromField(Field[] fields, string strField, string strPickListName)
        {
            string strRet = string.Empty;

            Field fldFieldToFind = FindField(fields, strField);
            if (fldFieldToFind == null)
            {
                throw new Exception("Could not get the " + strField + " field from the collection");
            }

            PickListValue plvValueToFind = FindPickListLabel(fldFieldToFind.PicklistValues, strPickListName);
            if (plvValueToFind != null)
            {
                strRet = plvValueToFind.Value;
            }

            return strRet;
        }

        /// <summary>
        /// Returns the label of a picklist when the value is sent
        /// </summary>
        /// <param name="fields">entity fields</param>
        /// <param name="strField">picklick to choose from</param>
        /// <param name="strPickListValue">value ("id") of picklist</param>
        /// <returns>picklist label</returns>
        protected static string PickListLabelFromValue(Field[] fields, string strField, string strPickListValue)
        {
            string strRet = string.Empty;

            Field fldFieldToFind = FindField(fields, strField);
            if (fldFieldToFind == null)
            {
                throw new Exception("Could not get the " + strField + " field from the collection");
            }

            PickListValue plvValueToFind = FindPickListValue(fldFieldToFind.PicklistValues, strPickListValue);
            if (plvValueToFind != null)
            {
                strRet = plvValueToFind.Label;
            }

            return strRet;
        }

        /// <summary>
        /// Used to find a specific Field in an array based on the name
        /// </summary>
        /// <param name="field">array containing Fields to search from</param>
        /// <param name="name">contains the name of the Field to search for</param>
        /// <returns>Field match</returns>
        protected static Field FindField(Field[] field, string name)
        {
            return Array.Find(field, element => element.Name == name);
        }

        /// <summary>
        /// Used to find a specific value in a picklist
        /// </summary>
        /// <param name="pickListValue">array of PickListsValues to search from</param>
        /// <param name="name">contains the name of the PickListValue to search for</param>
        /// <returns>PickListValue match</returns>
        protected static PickListValue FindPickListLabel(PickListValue[] pickListValue, string name)
        {
            return Array.Find(pickListValue, element => element.Label == name);
        }

        /// <summary>
        /// Used to find a specific value in a picklist
        /// </summary>
        /// <param name="pickListValue">array of PickListsValues to search from</param>
        /// <param name="valueID">contains the value of the PickListValue to search for</param>
        /// <returns>PickListValue match</returns>
        protected static PickListValue FindPickListValue(PickListValue[] pickListValue, string valueID)
        {
            return Array.Find(pickListValue, element => element.Value == valueID);
        }
    }
}
