using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace ol.clean
{
    public partial class Cleaner
    {
        private void Cleaner_Load(object sender, RibbonUIEventArgs e)
        {
            drpPeriod.SelectedItem = drpPeriod.Items[1];
        }

        private void btnClean_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PerformClean();
            }
            catch(System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // todo
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddDomain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PerformAdd(false);
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddExact_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PerformAdd(true);
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLogFolder_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var homeDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                System.Diagnostics.Process.Start(Cleaner.EnsureLogFolder(homeDir));
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnManageRules_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // todo
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnFindDomain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PerformFind(false);
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnFindExact_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PerformFind(true);
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PerformClean()
        {
            var now = DateTime.Now;
            var settings = Settings.Load();
            var app = Globals.ThisAddIn.Application;

            var folder = app.ActiveExplorer().CurrentFolder;
            if (!IsFolderValid("Clean", folder, null, false))
                return;

            var allItems = folder.Items;
            var emails = new List<MatchedItem>();

            for(int i=1; i <= allItems.Count; i++)
            {
                var mailItem = allItems[i] as MailItem;
                if (mailItem != null && mailItem.SentOn != null)
                {
                    var emailAddress = GetSenderSMTPAddress(mailItem);
                    if (!string.IsNullOrWhiteSpace(emailAddress))
                    {
                        var sent = mailItem.SentOn;
                        var rule = settings.FindRule(emailAddress);

                        if (sent != DateTime.MinValue && sent != DateTime.MaxValue)
                        {
                            if (rule != null && rule.Period > 0 && sent.AddDays(rule.Period) < now)
                            {
                                emails.Add(new MatchedItem(mailItem, rule, emailAddress, sent));
                            }
                        }
                    }
                    mailItem.Close(OlInspectorClose.olDiscard);
                }
            }

            if(emails.Count == 0)
            {
                MessageBox.Show("Found no emails matching any rules.", "Clean", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            emails.Sort(delegate (MatchedItem r1, MatchedItem r2)
            {
                int compareVal = r1.rule.Criteria.CompareTo(r2.rule.Criteria);
                if (compareVal == 0)
                    compareVal = r1.sent.CompareTo(r2.sent);
                return compareVal;
            });

            var logFilename = Cleaner.GetLogFilename(chkCommit.Checked);
            using (var log = new StreamWriter(logFilename, true, Encoding.UTF8))
            {
                log.WriteLine("\"Rule\",\"Exact\",\"Period\",\"Sent\",\"From\",\"Subject\"");
                foreach (MatchedItem nextItem in emails)
                {
                    var sent = nextItem.sent.ToString("yyyy-MM-dd HH:mm:ss");
                    var subj = string.IsNullOrWhiteSpace(nextItem.mail.Subject) ? "(blank)" : nextItem.mail.Subject.Replace("\"", "\"\"");
                    var txt1 = string.Format("\"{0}\",{1},{2},", nextItem.rule.Criteria, !nextItem.rule.EndsWith, nextItem.rule.Period);
                    var txt2 = string.Format("\"{0}\",\"{1}\",\"{2}\"", sent, nextItem.email, subj);
                    log.WriteLine(txt1 + txt2);

                    if (drpMarkRead.SelectedItem.Label == "Mark read")
                    {
                        nextItem.mail.UnRead = false;
                        nextItem.mail.Save();
                    }

                    if (drpMarkRead.SelectedItem.Label == "Mark unread")
                    {
                        nextItem.mail.UnRead = true;
                        nextItem.mail.Save();
                    }

                    if (chkCommit.Checked)
                        nextItem.mail.Delete();
                    else
                        nextItem.mail.Close(OlInspectorClose.olDiscard);
                }
            }

            if (chkCommit.Checked)
            {
                var msg = string.Format("Cleaned {0} emails.\r\n\r\nSee CSV file for a detailed log.", emails.Count.ToString());
                MessageBox.Show(msg, "Clean", MessageBoxButtons.OK, MessageBoxIcon.Information);
                drpMarkRead.SelectedItem = drpMarkRead.Items[0];
                chkCommit.Checked = false;
            }
            else
            {
                System.Diagnostics.Process.Start(logFilename);
            }
        }

        private void PerformAdd(bool addExact)
        {
            var app = Globals.ThisAddIn.Application;
            var explorer = app.ActiveExplorer();

            var folder = explorer.CurrentFolder;
            var allItems = explorer.Selection;
            if (!IsFolderValid("Add", folder, allItems, true))
                return;

            // Get list of email addresses.

            var emails = new List<string>();
            foreach (object nextItem in allItems)
            {
                var mailItem = nextItem as MailItem;
                if (mailItem != null)
                {
                    var emailAddress = GetSenderSMTPAddress(mailItem);
                    if (!string.IsNullOrWhiteSpace(emailAddress) && !emails.Any(a => string.Compare(emailAddress, a, StringComparison.InvariantCultureIgnoreCase) == 0))
                        emails.Add(emailAddress.ToLowerInvariant());
                }
            }

            // Perform a merge.

            var settings = Settings.Load();

            int iAddAcount = 0;
            foreach (var address in emails)
            {
                var rule = settings.Find(address, addExact);

                if (rule == null)
                {
                    if (addExact || !settings.IsExcluded(address))
                    {
                        settings.Add(address, addExact, 14);
                        iAddAcount++;
                    }
                }
            }

            if (iAddAcount > 0)
            {
                settings.Save();
                MessageBox.Show(string.Format("Inserted {0} items to the current list.", iAddAcount), "Add", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("No change was made to current list", "Add", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private string GetSenderSMTPAddress(MailItem mail)
        {
            // Method from https://msdn.microsoft.com/en-us/library/office/ff184624.aspx
            string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            if (mail == null)
                return null;

            if (mail.SenderEmailType == "EX")
            {
                var sender = mail.Sender;

                if (sender != null)
                {
                    // Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                        || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        // Use the ExchangeUser object PrimarySMTPAddress
                        var exchUser = sender.GetExchangeUser();
                        if (exchUser != null)
                            return exchUser.PrimarySmtpAddress;
                        else
                            return null;
                    }
                    else
                        return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                }
                else
                    return null;
            }
            else
            {
                var addr = mail.SenderEmailAddress;
                var indx = addr.IndexOf("=");
                return indx > 0 ? addr.Substring(indx + 1) : addr;
            }
        }

        private static string GetLogFilename(bool commit)
        {
            var homeDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var logDir = Cleaner.EnsureLogFolder(homeDir);

            var strCommit = commit ? "commit" : "logged";
            var name = string.Format("ol.clean {0} {1}.csv", DateTime.Now.ToString("yyyyMMdd HHmmssff"), strCommit);
            return Path.Combine(logDir, name);
        }

        private static string EnsureLogFolder(string homeDir)
        {
            var logDir = Path.Combine(homeDir, "ol.clean");
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);
            return logDir;
        }

        /// <summary>
        /// Perfomrs a search operation on the folder.
        /// https://msdn.microsoft.com/en-us/library/office/ff869309.aspx
        /// </summary>
        /// <param name="findExact">Exact match or domain search.</param>
        private void PerformFind(bool findExact)
        {
            var app = Globals.ThisAddIn.Application;
            var explorer = app.ActiveExplorer();

            var folder = explorer.CurrentFolder;
            var selection = explorer.Selection;

            if (selection == null || selection.Count == 0)
            {
                MessageBox.Show("No emails are selected", "Find", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (selection.Count != 1)
            {
                MessageBox.Show("Multiple selection is not allowed.", "Find", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var address = GetSenderSMTPAddress(selection[1]);
            if (string.IsNullOrWhiteSpace(address))
            {
                MessageBox.Show("No email address was found.", "Find", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var index = address.IndexOf('@');
            var useAddress = (findExact == false && index > 0) ? address.Substring(index + 1) : address;
            explorer.Search(string.Format("from:(\"{0}\")", useAddress), OlSearchScope.olSearchScopeCurrentFolder);
        }

        private bool IsFolderValid(string funcName, MAPIFolder folder, Selection selection, bool requireSelection)
        {
            if (folder == null || folder.Name != "Inbox")
            {
                MessageBox.Show("Not 'Inbox'", funcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (requireSelection)
            {
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("No emails are selected", funcName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            return true;
        }
    }

    internal class MatchedItem
    {
        internal MailItem mail;
        internal TypeRule rule;
        internal string email;
        internal DateTime sent;

        internal MatchedItem(MailItem mail, TypeRule rule, string email, DateTime sent)
        {
            this.mail = mail;
            this.rule = rule;
            this.email = email;
            this.sent = sent;
        }
    }
}
