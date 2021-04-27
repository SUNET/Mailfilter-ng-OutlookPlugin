using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace HalonSpamreport
{
    [ComVisible(true)]
    public class OutlookPlugin : Office.IRibbonExtensibility
    {
        private static string DefaultSpamButtonText = "Spam";
        private static string DefaultHamButtonText = "Non-spam";
        private static string DefaultForwardButtonText = "Forward to support";
        private static string DefaultButtonGroupText = "Halon Spamreport";
        private static string RegistryForwardingButtonText = "ForwardingButtonText";
        private static string RegistrySpamButtonText = "SpamButtonText";
        private static string RegistryHamButtonText = "HamButtonText";
        private static string RegistryButtonGroupText = "ButtonGroupText";
        private static string RegistryForwardingAddress = "ForwardingAddress";
        private static string ShortSpamAction = "s";
        private static string ShortHamAction = "n";

        private Office.IRibbonUI ribbon;
        private static string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        private static string MimePropertySchema = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}";
        private static string MimeHeaderRefid = "X-Halon-RPD-Refid";
        private static string MimeHeaderSpamClassification = "X-Spam-Flag";

        internal string ForwardingButtonText { get; set; }
        internal string SpamButtonText { get; set; }
        internal string HamButtonText { get; set; }
        internal string ButtonGroupText { get; set; }

        public OutlookPlugin()
        {
        }

        /// <summary>
        /// Gets all the headers of a mailitem
        /// </summary>
        /// <param name="mailItem"></param>
        /// <returns></returns>
        private string GetAllHeaders(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            if (null != mailItem)
            {
                return (string)mailItem.PropertyAccessor.GetProperty(TransportMessageHeadersSchema);
            }

            return string.Empty;
        }

        /// <summary>
        /// Finds a single mime-header among all the headers
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        private string GetHeader(Microsoft.Office.Interop.Outlook.MailItem mailItem, string header)
        {
            string headerValue = string.Empty;

            string allHeaders = GetAllHeaders(mailItem);
            if (!string.IsNullOrWhiteSpace(allHeaders))
            {
                var match = System.Text.RegularExpressions.Regex.Match(allHeaders, string.Format(CultureInfo.CurrentCulture, "{0}:[\\s]*(.*)[\\s]*\r\n", header));
                if (null != match && match.Groups.Count > 1)
                {
                    headerValue = match.Groups[1].Value.Trim(new char[] { ' ', '<', '>' });
                }
            }

            return headerValue;
        }

        /// <summary>
        /// Adds a header to a mailitem
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="header"></param>
        /// <param name="value"></param>
        private void AddHeader(Microsoft.Office.Interop.Outlook._MailItem mailItem, string header, string value)
        {
            if (null != mailItem)
            {
                mailItem.PropertyAccessor.SetProperty(string.Format("{0}/{1}", MimePropertySchema, header), value);
            }
        }

        private void ReportMail(string action, string shortAction)
        {
            List<Tuple<string, string, Microsoft.Office.Interop.Outlook.MailItem>> headerItems = new List<Tuple<string, string, Microsoft.Office.Interop.Outlook.MailItem>>();
            int failedMessages = 0;

            if (null == Globals.ThisAddIn.ApiUrl || string.IsNullOrEmpty(Globals.ThisAddIn.ApiUrl) ||
                null == Globals.ThisAddIn.ApiUser || string.IsNullOrEmpty(Globals.ThisAddIn.ApiUser) ||
                null == Globals.ThisAddIn.ApiPassword || string.IsNullOrEmpty(Globals.ThisAddIn.ApiPassword))
            {
                MessageBox.Show("Plugin is not configured.", "Error");
                return;
            }

            try
            {
                var errorMessage = string.Empty;
                var numberWithoutRefId = 0;
                var numberOfSelectedMessages = Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count;
                using (var wc = new WebClientEx(Globals.ThisAddIn.ApiUser, Globals.ThisAddIn.ApiPassword))
                {
                    foreach (Microsoft.Office.Interop.Outlook.MailItem mailitem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                    {
                        string queryData = string.Empty;

                        // Get (possible) spam classification header
                        var spamClassified = GetHeader(mailitem, MimeHeaderSpamClassification);

                        // Check if requested operation is appropriate
                        if (SpamButtonText == action && "YES" == spamClassified)
                        {
                            if (errorMessage == string.Empty && numberOfSelectedMessages == 1)
                            {
                                errorMessage = "This message is already spam-classified.";
                            }
                            continue; // Can't spam-classify mail thats already spam-classified
                        }
                        else if (HamButtonText == action && "YES" != spamClassified)
                        {
                            if (errorMessage == string.Empty && numberOfSelectedMessages == 1)
                            {
                                errorMessage = "This message is not spam-classified.";
                            }
                            continue; // Can't ham-classify mail thats not spam-classified
                        }

                        var refId = GetHeader(mailitem, MimeHeaderRefid);
                        if (refId == string.Empty)
                        {
                            if (errorMessage == string.Empty && numberOfSelectedMessages == 1)
                            {
                                errorMessage = "This message could not be reported due to missing mime-headers.";
                            }
                            numberWithoutRefId++;
                            continue;
                        }
                        var jsonData = string.Empty;
                        if (SpamButtonText == action)
                        {
                            var headersAndBody = GetAllHeaders(mailitem) + mailitem.Body;
                            jsonData = string.Format("{{\"type\":\"fn\",\"direction\":\"inbound\",\"ref_id\":\"{0}\",\"message\":\"{1}\"}}", refId, System.Convert.ToBase64String(Encoding.UTF8.GetBytes(headersAndBody)));
                        }
                        else if (HamButtonText == action)
                        {
                            jsonData = string.Format("{{\"type\":\"fp\",\"direction\":\"inbound\",\"ref_id\":\"{0}\",\"message\":\"\"}}", refId);
                        }

                        if (!string.IsNullOrEmpty(jsonData))
                        {
                            headerItems.Add(new Tuple<string, string, Microsoft.Office.Interop.Outlook.MailItem>(refId, jsonData, SpamButtonText == action ? mailitem : null));
                        }
                    }

                    // No messages were selected
                    if (headerItems.Count <= 0)
                    {
                        MessageBox.Show(errorMessage == string.Empty ? "No messages could be reported." : errorMessage);
                        return;
                    }

                    var apiUri = new Uri(Globals.ThisAddIn.ApiUrl);
                    foreach (var reportTuple in headerItems)
                    {
                        try
                        {
                            wc.UploadString(apiUri, reportTuple.Item2);
                            Globals.ThisAddIn.LogMessage(string.Format("Voting {0} (authenticated): {1}", action, reportTuple.Item1), string.Empty, System.Diagnostics.EventLogEntryType.Information);
                        }
                        catch (WebException we)
                        {
                            if (we.Response.ContentLength > 0)
                            {
                                var body = (new StreamReader(we.Response.GetResponseStream()).ReadToEnd());
                                Globals.ThisAddIn.LogMessage(body, string.Empty, System.Diagnostics.EventLogEntryType.Error);
                            }
                            // Ignore 404 exceptions (spam entry not found)
                            if (we.Response is HttpWebResponse && (we.Response as HttpWebResponse).StatusCode == HttpStatusCode.NotFound)
                            {
                                failedMessages++;
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }

                    foreach (var reportTuple in headerItems)
                    {
                        if (null != reportTuple.Item3)
                        {
                            try
                            {
                                if (null != Globals.ThisAddIn.JunkFolder)
                                {
                                    reportTuple.Item3.Move(Globals.ThisAddIn.JunkFolder);
                                }
                            }
                            catch (Exception)
                            {
                                // Quench error and continue
                            }
                        }
                    }

                    if (failedMessages > 0 && failedMessages < headerItems.Count)
                    {
                        var msg = "Some of the messages could not be reported on";
                        if (numberWithoutRefId > 0)
                        {
                            msg += " (one or more are missing Halon mime-headers)";
                        }
                        msg += ".";
                        MessageBox.Show(msg);
                    }
                    else if (failedMessages > 0)
                    {
                        var msg = "None of the messages could not be reported on";
                        if (numberWithoutRefId > 0)
                        {
                            msg += " (one or more are missing Halon mime-headers)";
                        }
                        msg += ".";
                        MessageBox.Show(msg);
                    }
                    else
                    {
                        MessageBox.Show(string.Format("The selected message(s) has been reported as {0}", action));
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogMessage(ex.Message, ex.StackTrace);
                if (Globals.ThisAddIn.ShowPopups)
                {
                    MessageBox.Show("An error occurred, check the eventlog for detailed information.");
                }
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            // Only add ribbon to the main window, not in individual mails
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                var xmlData = GetResourceText("HalonSpamreport.OutlookPlugin.xml");
                // Workaround, patching the resouce XML with the configured registry texts since Outlook 2010 does not fire the getLabel events properly.
                // If Outlook 2010 compatibility is not required, the xml should define a getLabel callback and provide the texts through that instead.
                xmlData = xmlData.Replace("GroupPlaceholderLabel", ThisAddIn.LoadResourceString(RegistryButtonGroupText, DefaultButtonGroupText));
                this.SpamButtonText = ThisAddIn.LoadResourceString(RegistrySpamButtonText, DefaultSpamButtonText);
                xmlData = xmlData.Replace("SpamButtonPlaceholderLabel", this.SpamButtonText);
                this.HamButtonText = ThisAddIn.LoadResourceString(RegistryHamButtonText, DefaultHamButtonText);
                xmlData = xmlData.Replace("HamButtonPlaceholderLabel", this.HamButtonText);

                // Remove forwarding button if a forwarding address has not been configured in the registry, otherwise set button text
                var forwardingAddress = ThisAddIn.LoadResourceString(RegistryForwardingAddress, string.Empty);
                var forwardingButtonText = ThisAddIn.LoadResourceString(RegistryForwardingButtonText, DefaultForwardButtonText);
                if (string.Empty == forwardingAddress || string.Empty == forwardingButtonText)
                {
                    // Find button definition in xml
                    var buttonMatch = Regex.Match(xmlData, "<.*\"ForwardButton\".*/>");
                    if (buttonMatch.Groups.Count > 0)
                    {
                        xmlData = xmlData.Replace(buttonMatch.Groups[0].ToString(), string.Empty);
                    }
                }
                else
                {
                    xmlData = xmlData.Replace("ForwardButtonPlaceholderLabel", forwardingButtonText);
                }

                return xmlData;
            }

            return null;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public System.Drawing.Bitmap getOutlookPluginImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "SpamButton":
                    return HalonSpamreport.Properties.Resources.spam_icon;
                case "HamButton":
                    return HalonSpamreport.Properties.Resources.ham_icon;
                case "ForwardButton":
                    return HalonSpamreport.Properties.Resources.forward_icon;
            }

            return null;
        }

        public void spam_Click(IRibbonControl control)
        {
            this.ReportMail(SpamButtonText, ShortSpamAction);
        }

        public void ham_Click(IRibbonControl control)
        {
            this.ReportMail(HamButtonText, ShortHamAction);
        }

        public void forward_Click(IRibbonControl control)
        {
            if (null == Globals.ThisAddIn.ForwardingSubject || string.IsNullOrEmpty(Globals.ThisAddIn.ForwardingSubject) ||
                null == Globals.ThisAddIn.ForwardingBody || string.IsNullOrEmpty(Globals.ThisAddIn.ForwardingBody))
            {
                MessageBox.Show("Plugin is not configured.", "Error");
                return;
            }

            try
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count <= 0)
                {
                    return;
                }

                if (MessageBox.Show("Send selected mail?", "Send", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                Microsoft.Office.Interop.Outlook._MailItem mi = (Microsoft.Office.Interop.Outlook._MailItem)Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mi.Subject = Globals.ThisAddIn.ForwardingSubject;
                mi.Body = Globals.ThisAddIn.ForwardingBody;
                mi.Recipients.Add(Globals.ThisAddIn.ForwardingAddress);
                foreach (Microsoft.Office.Interop.Outlook.MailItem mailitem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                {
                    while (mailitem.Attachments.Count > 0)
                    {
                        mailitem.Attachments.Remove(1);
                    }

                    mi.Attachments.Add(mailitem);
                }

                if (null != Globals.ThisAddIn.ForwardingMimeHeader && !string.IsNullOrEmpty(Globals.ThisAddIn.ForwardingMimeHeader) &&
                    null != Globals.ThisAddIn.ForwardingMimeValue && !string.IsNullOrEmpty(Globals.ThisAddIn.ForwardingMimeValue))
                {
                    AddHeader(mi, Globals.ThisAddIn.ForwardingMimeHeader, Globals.ThisAddIn.ForwardingMimeValue);
                }

                mi.Send();

                MessageBox.Show("The message(s) has been forwarded.");
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogMessage(ex.Message, ex.StackTrace);
                if (Globals.ThisAddIn.ShowPopups)
                {
                    MessageBox.Show("An error occurred, check the eventlog for detailed information.");
                }
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
