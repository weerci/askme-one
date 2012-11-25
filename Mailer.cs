using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;

namespace Askme
{
    class Mailer
    {
        // Public 
        public void SendMail(TypeGetReport typeGet, bool archiveMail)
        {
            this.typeGet = typeGet;
            this.archiveMail = archiveMail;
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.Credentials = new NetworkCredential(smtpUserName, smtpUserPass);
            
            MailMessage message = new MailMessage(msgFrom, msgTo, msgSubject, msgBody);
            GetAttachment(message);

            try
            {
                client.Send(message);
                Log.ToLog(CommonString.EmailSended);
            }
            catch (SmtpException err)
            {
                Log.ToLog(err.Message);
                throw;
            }
        }
        public void LoadData()
        {
            smtpHost = Common.SysProp.SmtpServer;
            smtpPort = Common.SysProp.Port;
            smtpUserName = Common.SysProp.login;
            smtpUserPass = Common.SysProp.password;
            msgFrom = Common.SysProp.MessageFrom;
            msgTo = Common.SysProp.MessageTo;
            msgSubject = Common.SysProp.MessageSubject;
            msgBody = Common.SysProp.MessageBody;
        }

        // Private
        private void GetAttachment(MailMessage message)
        {
            try
            {
                Attachment data;
                if (this.archiveMail)
                {
                    ZipResultFile();
                    data = new Attachment(Common.XML_ZIP_RESULT(this.typeGet), MediaTypeNames.Application.Octet);
                }
                else
                    data = new Attachment(Common.XML_RESULT(this.typeGet), MediaTypeNames.Application.Octet);
                
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(Common.XML_RESULT(this.typeGet));
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(Common.XML_RESULT(this.typeGet));
                disposition.ReadDate = System.IO.File.GetLastAccessTime(Common.XML_RESULT(this.typeGet));
                message.Attachments.Add(data);
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }
        }
        private void ZipResultFile()
        {
            try
            {
                using (ZipOutputStream s = new ZipOutputStream(File.Create(Common.XML_ZIP_RESULT(this.typeGet)))) 
                {
                    s.SetLevel(5); // 0 - store only to 9 - means best compression
                    byte[] buffer = new byte[4096];
                    ZipEntry entry = new ZipEntry(Path.GetFileName(Common.XML_RESULT(this.typeGet)));

                    entry.DateTime = DateTime.Now;
                    s.PutNextEntry(entry);
                    using (FileStream fs = File.OpenRead(Common.XML_RESULT(this.typeGet)))
                    {
                        int sourceBytes;
                        do
                        {
                            sourceBytes = fs.Read(buffer, 0, buffer.Length);
                            s.Write(buffer, 0, sourceBytes);
                        } while (sourceBytes > 0);
                    }
                }
                File.Delete(Common.XML_RESULT(this.typeGet)); 
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }
        }

        // Fields
        private String smtpHost = "SMTP.SERVER.RU";
        private int smtpPort = 25;
        private String smtpUserName = "LOGIN";
        private String smtpUserPass = "PASSWORD";
        private String msgFrom = "LOGIN@SERVER.RU";
        private String msgTo = "KUDA@TO.RU";
        private String msgSubject = "Тема письма";
        private String msgBody = "Тело письма";
        private TypeGetReport typeGet;
        private bool archiveMail;
    }
}
