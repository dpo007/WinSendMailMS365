# WinSendMailMS365
Basic [Sendmail](https://linux.die.net/man/8/sendmail.sendmail) replacement for Windows, written in C#.

Takes raw email fed via Console/StdIn, and sends it via MS365's Graph API using modern authentication methods.

Note: Supports plain-text bodies only.

My intended usage is with IIS, PHP and a setup of Drupal (all on-prem).

Based off my previous WinSendMail SMTP sender.

* Depends on MimeKitLite, Microsoft Graph, and Azure Identity packages (and their dependencies).
* Currently targets .Net Framework 4.7.2.
* Requires creating an Azure AD app registration.

