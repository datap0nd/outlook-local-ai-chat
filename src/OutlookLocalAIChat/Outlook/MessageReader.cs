using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class MessageReader
    {
        private readonly object _outlookApplication;

        public MessageReader(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public MessageSnapshot CaptureCurrent()
        {
            object item = null;
            object inspector = null;
            object explorer = null;
            object selection = null;

            try
            {
                dynamic application = _outlookApplication;
                inspector = application.ActiveInspector();
                if (inspector != null)
                {
                    dynamic activeInspector = inspector;
                    item = activeInspector.CurrentItem;
                }

                if (item == null)
                {
                    explorer = application.ActiveExplorer();
                    if (explorer != null)
                    {
                        dynamic activeExplorer = explorer;
                        selection = activeExplorer.Selection;
                        dynamic currentSelection = selection;
                        if (currentSelection != null && currentSelection.Count > 0)
                        {
                            item = currentSelection.Item(1);
                        }
                    }
                }

                if (item == null)
                {
                    throw new InvalidOperationException(
                        "Select or open an email in Outlook first.");
                }

                dynamic mail = item;
                var messageClass = SafeString(() => mail.MessageClass);
                if (!messageClass.StartsWith(
                    "IPM.Note",
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "The selected Outlook item is not an email.");
                }

                return CaptureItem(item);
            }
            finally
            {
                Release(selection);
                Release(explorer);
                Release(inspector);
                Release(item);
            }
        }

        public MessageSnapshot CaptureById(
            string entryId,
            string storeId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                throw new ArgumentException(
                    "A message entry ID is required.",
                    nameof(entryId));
            }

            object session = null;
            object item = null;
            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                item = string.IsNullOrWhiteSpace(storeId)
                    ? outlookSession.GetItemFromID(entryId)
                    : outlookSession.GetItemFromID(entryId, storeId);
                return CaptureItem(item);
            }
            finally
            {
                Release(item);
                Release(session);
            }
        }

        public IReadOnlyList<MessageSnapshot> CaptureActiveSelectionMany()
        {
            object explorer = null;
            object selection = null;
            try
            {
                dynamic application = _outlookApplication;
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    throw new InvalidOperationException(
                        "Open an Outlook mailbox view and select one to five emails.");
                }

                dynamic activeExplorer = explorer;
                selection = activeExplorer.Selection;
                return CaptureSelectionMany(selection);
            }
            finally
            {
                Release(selection);
                Release(explorer);
            }
        }

        public MessageSnapshot CaptureSelection(object selection)
        {
            var messages = CaptureSelectionMany(selection);
            if (messages.Count != 1)
            {
                throw new InvalidOperationException(
                    "Select exactly one email before using this action.");
            }

            return messages[0];
        }

        public IReadOnlyList<MessageSnapshot> CaptureSelectionMany(
            object context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            object selection = context;
            var releaseSelection = false;
            try
            {
                int count;
                try
                {
                    dynamic selectedItems = selection;
                    count = Convert.ToInt32(selectedItems.Count);
                }
                catch
                {
                    dynamic window = context;
                    selection = window.Selection;
                    releaseSelection = !ReferenceEquals(
                        selection,
                        context);
                    dynamic selectedItems = selection;
                    count = Convert.ToInt32(selectedItems.Count);
                }

                if (count < 1)
                {
                    throw new InvalidOperationException(
                        "Select at least one email before using Send to MailAI.");
                }

                if (count > MailboxWorkingSet.MaxMessages)
                {
                    throw new InvalidOperationException(
                        "Select no more than five emails before using Send to MailAI.");
                }

                var messages = new List<MessageSnapshot>(count);
                dynamic items = selection;
                for (var index = 1; index <= count; index++)
                {
                    object item = null;
                    try
                    {
                        item = items.Item(index);
                        messages.Add(CaptureItem(item));
                    }
                    finally
                    {
                        Release(item);
                    }
                }

                return MailboxWorkingSet.Normalize(messages);
            }
            finally
            {
                if (releaseSelection)
                {
                    Release(selection);
                }
            }
        }

        internal static MessageSnapshot CaptureItem(object item)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            object parent = null;
            try
            {
                dynamic mail = item;
                var messageClass = SafeString(() => mail.MessageClass);
                if (!messageClass.StartsWith(
                    "IPM.Note",
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "The Outlook item is not an email.");
                }

                parent = SafeObject(() => mail.Parent);
                var storeId = string.Empty;
                if (parent != null)
                {
                    dynamic folder = parent;
                    storeId = SafeString(() => folder.StoreID);
                }

                var receivedAt =
                    SafeDateTime(() => mail.ReceivedTime) ??
                    SafeDateTime(() => mail.SentOn) ??
                    SafeDateTime(() => mail.CreationTime);

                return new MessageSnapshot(
                    SafeString(() => mail.EntryID),
                    storeId,
                    TextBoundary.PlainText(
                        SafeString(() => mail.Subject),
                        1000),
                    TextBoundary.PlainText(
                        BuildSender(mail),
                        1000),
                    TextBoundary.PlainText(
                        SafeString(() => mail.To),
                        2000),
                    receivedAt,
                    TextBoundary.PlainText(
                        SafeString(() => mail.Body),
                        TextBoundary.MaxMessageBodyCharacters));
            }
            finally
            {
                Release(parent);
            }
        }

        internal static bool IsMailItem(object item)
        {
            if (item == null)
            {
                return false;
            }

            dynamic mail = item;
            return SafeString(() => mail.MessageClass).StartsWith(
                "IPM.Note",
                StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSender(dynamic mail)
        {
            var name = SafeString(() => mail.SenderName);
            var address = SafeString(() => mail.SenderEmailAddress);
            if (name.Length == 0)
            {
                return address;
            }

            if (address.Length == 0 ||
                name.Equals(address, StringComparison.OrdinalIgnoreCase))
            {
                return name;
            }

            return name + " <" + address + ">";
        }

        private static object SafeObject(Func<object> reader)
        {
            try
            {
                return reader();
            }
            catch
            {
                return null;
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime? SafeDateTime(Func<object> reader)
        {
            try
            {
                var value = reader();
                if (value is DateTime)
                {
                    return (DateTime)value;
                }
            }
            catch
            {
            }

            return null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
