Halon Spamreport v1.0
=====================

Halon Spamreport is an Outlook plugin that allows a user to send feedback to the spamfilter.
It exposes three new buttons in the regular Outlook GUI, "Spam", "Non-spam" and "Forward to support".
Spam and Non-spam reports the currently selected mails to the spamfilter, whereas Forward forwards the
mail to a preconfigured e-mailaddress.

Requirements
The plugin requires the .NET Framework version 4.5 or greater.

Configuration
By default, the Spamreportplugin doesn't need any configuration if MIME-headers with voting links are present in the email messages AND unauthenicated voting is allowed.

If you want to customize the configuration of the Spamreportplugin, it's done from the windows registry.
Base registry path: HKEY_LOCAL_MACHINE\Software\SUNET\HalonSpamreport
(Individual keys can be overridden in HKEY_CURRENT_USER\Software\SUNET\HalonSpamreport)

The following registry keys needs to be configured (User/Password will be sent as Basic Authentication):

ApiUrl			string		Url of Halon API services [required]
ApiUser			string		User for Halon API services [required]
ApiPassword		string		Password for Halon API services [required]

If you want an additional button for forwarding the marked email(s) as attachments and sent to email address of your choosing, the following keys can be configured:

ForwardingAddress	string		Address to forward mail to [required]
ForwardingButtonText	string		Text of forwarding button [optional]
ForwardingSubject	string		Text of mail subject if forwarding mail [optional]
ForwardingBody		string		Text of mail body if forwarding mail [optional]
ForwardingMimeHeader	string		Mime header to add to forwarded mail [optional]
ForwardingMimeValue	string		Mime header value to add to forwarded mail [optional]

Other customizations that optionally can be configured:

SpamButtonText		string		Text of spam button [optional]
HamButtonText		string		Text of non-spam button [optional]
ButtonGroupText		string		Text of button toolbar [optional]
ShowPopups		bool		Confirmation popup message then voting [optional]

Logging
The plugin logs errors to the Windows EventLog under the name "HalonSpamreport Plugin".
