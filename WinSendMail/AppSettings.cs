namespace WinSendMailMS365
{
    /// <summary>
    /// Class containing settings for application.
    /// </summary>
    internal class AppSettings
    {
        public string MS365ClientID { get; set; }
        public string MS365ClientSecret { get; set; }
        public string MS365TenantID { get; set; }
        public string MS365SendingUser { get; set; }
        public bool HTMLDecodeContent { get; set; }
        public bool RemoveDuplicateBlankLines { get; set; }
        public bool SaveEmailsToDisk { get; set; }
    }
}