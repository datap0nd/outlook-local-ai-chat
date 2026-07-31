using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    internal sealed class DraftService
    {
        private readonly object _outlookApplication;

        internal DraftService(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        internal DraftSession CreateReplyDraft(
            MessageSnapshot source,
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            if (source == null || !source.CanReply)
            {
                throw new InvalidOperationException(
                    "The resolved source email is not available for reply.");
            }

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
                var quotedHtml = SafeString(
                    () => replyMail.HTMLBody);
                var draft = new DraftSession(
                    reply,
                    "reply",
                    quotedHtml);
                draft.Update(
                    body,
                    boldPhrases,
                    null,
                    null,
                    null);
                reply = null;
                return draft;
            }
            finally
            {
                Release(reply);
                Release(original);
                Release(session);
            }
        }

        internal DraftSession CreateNewDraft(
            string body,
            IReadOnlyList<string> boldPhrases,
            string subject,
            string to,
            string cc)
        {
            object draftItem = null;
            try
            {
                dynamic application = _outlookApplication;
                draftItem = application.CreateItem(0);
                var draft = new DraftSession(
                    draftItem,
                    "new",
                    string.Empty);
                draft.Update(
                    body,
                    boldPhrases,
                    subject,
                    to,
                    cc);
                draftItem = null;
                return draft;
            }
            finally
            {
                Release(draftItem);
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

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }

    internal sealed class DraftSession : IDisposable
    {
        private object _mailItem;
        private readonly string _kind;
        private readonly string _quotedHtml;
        private string _body = string.Empty;

        internal DraftSession(
            object mailItem,
            string kind,
            string quotedHtml)
        {
            _mailItem = mailItem ??
                throw new ArgumentNullException(nameof(mailItem));
            _kind = kind ?? string.Empty;
            _quotedHtml = quotedHtml ?? string.Empty;
        }

        internal DraftReference Reference
        {
            get
            {
                EnsureAvailable();
                dynamic mail = _mailItem;
                return new DraftReference(
                    _kind,
                    SafeString(() => mail.Subject),
                    SafeString(() => mail.To),
                    SafeString(() => mail.CC),
                    _body);
            }
        }

        internal void Update(
            string body,
            IReadOnlyList<string> boldPhrases,
            string subject,
            string to,
            string cc)
        {
            EnsureAvailable();
            var boundedBody = TextBoundary.PlainText(
                body,
                TextBoundary.MaxAssistantCharacters);
            var content = SafeDraftHtml.FormatContent(
                boundedBody,
                boldPhrases);
            dynamic mail = _mailItem;

            if (subject != null)
            {
                var formattedSubject = SafeModelText.Format(
                    subject,
                    255);
                mail.Subject = TextBoundary.SingleLine(
                    formattedSubject.PlainText,
                    255);
            }

            if (to != null)
            {
                mail.To = TextBoundary.SingleLine(to, 2000);
            }

            if (cc != null)
            {
                mail.CC = TextBoundary.SingleLine(cc, 2000);
            }

            mail.HTMLBody = _kind == "reply" &&
                _quotedHtml.Length > 0
                ? content.Html + "<br><br>" + _quotedHtml
                : content.Html;
            mail.Save();
            mail.Display(false);
            _body = content.PlainText;
        }

        public void Dispose()
        {
            var item = _mailItem;
            _mailItem = null;
            if (item != null && Marshal.IsComObject(item))
            {
                Marshal.ReleaseComObject(item);
            }
        }

        private void EnsureAvailable()
        {
            if (_mailItem == null)
            {
                throw new InvalidOperationException(
                    "The linked Outlook draft is no longer available.");
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
    }
}
