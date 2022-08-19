namespace WinSendMailMS365
{
    /// <summary>
    /// Class containing settings for application.
    /// </summary>
    internal class AppSettings
    {
        public string MS365ClientID { get; internal set; }
        public string MS365ClientSecret { get; internal set; }
        public string MS365TenantID { get; internal set; }
        public string MS365SendingUser { get; internal set; }
        public bool HTMLDecodeContent { get; internal set; }
        public bool SaveEmailsToDisk { get; internal set; }
    }
}