using System;

namespace Graph2AutoTask
{
    public class AutomationConfig
    {
        public MailboxConfig[] MailBoxes { get; set; }
    }
    public class AutotaskConfig
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public AutotaskDefaultsConfig Defaults { get; set; }
    }
    public class GraphConfig
    {
        public string TenantID { get; set; }
        public string ClientID { get; set; }
        public string ClientSecret { get; set; }
        internal string ClientAuthority => $"https://login.microsoftonline.com/{TenantID}/v2.0";
    }

    public class MailboxConfig
    {
        public string MailBox { get; set; }
        public MailboxProcessingConfig Processing { get; set; }
        public MailboxFolderConfig Folders { get; set; }
        public AutotaskConfig Autotask { get; set; }
        public GraphConfig Graph { get; set; }
    }
    public class MailboxProcessingConfig
    {
        public bool Enabled { get; set; }
        public TimeSpan CheckDelay { get; set; }
        public bool UnreadOnly { get; set; }
    }
    public class MailboxFolderConfig
    {
        public string Incoming { get; set; }
        public string Processed { get; set; }
        public string Failed { get; set; }
    }
    public class AutotaskDefaultsConfig
    {
        public AutotaskTicketDefaultsConfig Ticket { get; set; }
        public AutotaskTicketNoteDefaultsConfig TicketNote { get; set; }
        public AutotaskAttachmentDefaultsConfig Attachment { get; set; }
    }
    public class AutotaskTicketDefaultsConfig
    {
        public string Account { get; set; }
        public string Status { get; set; }
        public string Priority { get; set; }
        public string Queue { get; set; }
        public string Source { get; set; }
        public string WorkType { get; set; }
        public TimeSpan DueDateOffset { get; set; }
    }
    public class AutotaskTicketNoteDefaultsConfig
    {
        public string Type { get; set; }
        public string Publish { get; set; }
    }

    public class AutotaskAttachmentDefaultsConfig
    {

        public bool ResizeLargeImages { get; set; }
        public bool CompressLargeItems { get; set; }

        public AutotaskAttachmentDefaultsConfigImage Image {get; set;}
        public string Publish { get; set; }
    }
    public class AutotaskAttachmentDefaultsConfigImage
    {
        public int MinWidth { get; set; }
        public int MinHeight { get; set; }
    }
}
