using System;
using System.Runtime.InteropServices;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class DraftService
    {
        private readonly object _outlookApplication;

        public DraftService(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public void CreateReplyDraft(
            MessageSnapshot source,
            string draftText)
        {
            if (source == null || !source.CanReply)
            {
                throw new InvalidOperationException(
                    "Refresh the selected email before creating a reply.");
            }

            var boundedText = RequireDraftText(draftText);
            object session = null;
            object original = null;
            object reply = null;

            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                original = source.StoreId.Length > 0
                    ? outlookSession.GetItemFromID(
                        source.EntryId,
                        source.StoreId)
                    : outlookSession.GetItemFromID(source.EntryId);
                dynamic originalMail = original;
                reply = originalMail.Reply();
                dynamic replyMail = reply;
                var quotedBody = Convert.ToString(replyMail.Body) ?? string.Empty;
                replyMail.Body =
                    boundedText + Environment.NewLine +
                    Environment.NewLine + quotedBody;
                replyMail.Save();
                replyMail.Display(false);
            }
            finally
            {
                Release(reply);
                Release(original);
                Release(session);
            }
        }

        public void CreateNewDraft(
            string draftText,
            string subject = null,
            string to = null,
            string cc = null)
        {
            var boundedText = RequireDraftText(draftText);
            var boundedSubject = TextBoundary.SingleLine(
                subject,
                255);
            var boundedTo = TextBoundary.SingleLine(
                to,
                2000);
            var boundedCc = TextBoundary.SingleLine(
                cc,
                2000);
            object draft = null;

            try
            {
                dynamic application = _outlookApplication;
                draft = application.CreateItem(0);
                dynamic mail = draft;
                mail.Subject = boundedSubject;
                mail.To = boundedTo;
                mail.CC = boundedCc;
                mail.Body = boundedText;
                mail.Save();
                mail.Display(false);
            }
            finally
            {
                Release(draft);
            }
        }

        private static string RequireDraftText(string value)
        {
            var bounded = TextBoundary.PlainText(
                value,
                TextBoundary.MaxAssistantCharacters);
            if (bounded.Length == 0)
            {
                throw new InvalidOperationException(
                    "Ask the assistant for draft text first.");
            }

            return bounded;
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
