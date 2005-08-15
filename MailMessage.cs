#define OPENSMTP
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

#if MAILMILL
using MailMill;
#elif OPENSMTP
using OpenSmtp.Mail;
#endif

namespace MapiToMaildir {
    public enum AddressType {
        To,
        Cc,
        Bcc
    }

    class MailMessage {
#if MAILMILL
        private Mail _msg;
 
        public MailMessage() {
            _msg = new Mail();
        }

        public string From {
            get {
                return _msg.From;
            }

            set {
                _msg.From = value;
            }
        }

        public DateTime Date {
            get {
                String dt = (String)_msg.Headers["Date"];
                if (dt == null) {
                    return DateTime.MinValue;
                }

                return DateTime.Parse(dt);
            }

            set {
                _msg.Headers["Date"] = value.ToUniversalTime().ToString("R");
            }
        }

        public string ReplyTo {
            get {
                return _msg.ReplyTo.ToString();
            }

            set {
                _msg.ReplyTo.Add(value);
            }
        }

        public string Subject {
            get {
                return _msg.Subject;
            }

            set {
                _msg.Subject = value;
            }
        }

        public string Body {
            get {
                return _msg.TextBody;
            }

            set {
                _msg.TextBody = value;
            }
        }

        public string HtmlBody {
            get {
                return _msg.HtmlBody;
            }

            set {
                _msg.HtmlBody = value;
            }
        }

        public static String CreateEmailAddress(String addr, String name) {
            return Mail.CreateEmailAddress(addr, name);
        }

        public void AddCustomHeader(String hdr, String body) {
            _msg.Headers.Add(hdr, body);
        }

        public void AddRecipient(String address, String name, AddressType type) {
            switch (type) {
                case AddressType.To:
                    _msg.Tos.Add(MailMessage.CreateEmailAddress(address, name));
                    break;
                case AddressType.Cc:
                    _msg.Ccs.Add(MailMessage.CreateEmailAddress(address, name));
                    break;
                case AddressType.Bcc:
                    _msg.Bccs.Add(MailMessage.CreateEmailAddress(address, name));
                    break;
            }
        }

        public void AddAttachment(Stream data, String filename, String mimeType) {
            MailAttachment at = new MailAttachment();
            at.Name = filename;
            at.Type = mimeType;

            using (BinaryReader rdr = new BinaryReader(data)) {
                at.Data = rdr.ReadBytes((int)data.Length);
            }

            _msg.Attachments.Add(at);
        }

        public void Save(String filename) {
            _msg.Save(filename);
        }

#endif

#if OPENSMTP
        public class EmailAddress {
            private OpenSmtp.Mail.EmailAddress _addr;

            public EmailAddress(String address, String name) {
                _addr = new OpenSmtp.Mail.EmailAddress(address, name);
            }

            internal EmailAddress(OpenSmtp.Mail.EmailAddress addr) {
                _addr = addr;
            }

            internal OpenSmtp.Mail.EmailAddress InternalAddress {
                get {
                    return _addr;
                }
            }
        }
        private OpenSmtp.Mail.MailMessage _msg;
 
        public MailMessage() {
            _msg = new OpenSmtp.Mail.MailMessage();
        }

        public EmailAddress From {
            get {
                return new EmailAddress(_msg.From);
            }

            set {
                _msg.From = value.InternalAddress;
            }
        }

        public DateTime Date {
            get {
                return _msg.Date;
            }

            set {
                _msg.Date = value;
            }
        }

        public EmailAddress ReplyTo {
            get {
                return new EmailAddress(_msg.ReplyTo);
            }

            set {
                _msg.ReplyTo = value.InternalAddress;
            }
        }

        public string Subject {
            get {
                return _msg.Subject;
            }

            set {
                _msg.Subject = value;
            }
        }

        public string Body {
            get {
                return _msg.Body;
            }

            set {
                _msg.Body = value;
            }
        }

        public string HtmlBody {
            get {
                return _msg.HtmlBody;
            }

            set {
                _msg.HtmlBody = value;
            }
        }

        public static EmailAddress CreateEmailAddress(String addr, String name) {
            return new EmailAddress(addr, name);
        }

        public void AddCustomHeader(String hdr, String body) {
            _msg.AddCustomHeader(hdr, body);
        }

        public void AddRecipient(String address, String name, AddressType type) {
            switch (type) {
                case AddressType.To:
                    _msg.AddRecipient(new OpenSmtp.Mail.EmailAddress(address, name), OpenSmtp.Mail.AddressType.To);
                    break;
                case AddressType.Cc:
                    _msg.AddRecipient(new OpenSmtp.Mail.EmailAddress(address, name), OpenSmtp.Mail.AddressType.Cc);
                    break;
                case AddressType.Bcc:
                    _msg.AddRecipient(new OpenSmtp.Mail.EmailAddress(address, name), OpenSmtp.Mail.AddressType.Bcc);
                    break;
            }
        }

        public void AddAttachment(Stream data, String filename, String mimeType, String contentId) {
            Attachment attach = new Attachment(data, filename);
            attach.MimeType = mimeType;
            attach.ContentId = contentId;

            _msg.AddAttachment(attach);
        }

        public void Save(String filename) {
            _msg.Save(filename);
        }
#endif
    }
}
