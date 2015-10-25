using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;

namespace KGSA
{
    class KgsaEmail
    {
        FormMain main;
        public DataTable emailDb;

        public KgsaEmail(FormMain form)
        {
            this.main = form;
            LoadEmailDb();
        }

        private void LoadEmailDb()
        {
            emailDb = main.database.GetSqlDataTable("SELECT * FROM tblEmail");
            Log.d("Hentet e-post adresser fra tblEmail. Antall linjer: " + emailDb.Rows.Count);
        }

        private string FilterString(string str, DateTime date)
        {
            try
            {
                if (str.Length == 0)
                    return "Ingen tekst";

                str = str.Replace("(ukedag)", date.ToString("dddd"));
                str = str.Replace("(dag)", date.ToString("d."));
                str = str.Replace("(måned)", date.ToString("MMMM"));
                return str;
            }
            catch
            {
                return "";
            }
        }

        public void ImportOldEmailList(string file)
        {
            try
            {
                if (File.Exists(file))
                {
                    string var = File.ReadAllText(FormMain.settingsPath + @"\emailRecipients.txt");
                    if (var.Length > 3)
                    {
                        var reg = new Regex("\".*?\"");
                        var matches = reg.Matches(var);
                        foreach (var item in matches)
                        {
                            string s = item.ToString();

                            var regg = new Regex("\\<(.*?)\\>");
                            string email = Convert.ToString(regg.Match(item.ToString()));
                            email = email.Substring(1, email.Length - 2);
                            string name = s.Substring(1, s.IndexOf("<") - 1);

                            if (Add(name, email, "Full", true))
                                Log.n("Importert epost: " + email + " - " + name, Color.Green);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public bool Add(string nameEmail, string addressEmail, string typeEmail, bool quickEmail)
        {
            nameEmail = nameEmail.Trim();

            if (!IsValidEmail(addressEmail))
            {
                Log.n("E-post adressen er i feil format.");
                return false;
            }

            if (CheckDuplicate(addressEmail))
                return false; // Denne adresse finnes allerede

            try
            {
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                using (SqlCeCommand cmd = new SqlCeCommand("insert into tblEmail(Name, Address, Type, Quick) values (@Val1, @val2, @val3, @val4)", con))
                {
                    cmd.Parameters.AddWithValue("@Val1", nameEmail);
                    cmd.Parameters.AddWithValue("@Val2", addressEmail);
                    cmd.Parameters.AddWithValue("@Val3", typeEmail);
                    cmd.Parameters.AddWithValue("@Val4", quickEmail);
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return false;
            }
        }

        public bool Remove(string id)
        {
            try
            {
                int result = 0;
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM tblEmail WHERE Id = " + id, con))
                {
                    result = com.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();
                if (result > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return false;
            }
        }

        private bool CheckDuplicate(string address)
        {
            var con = new SqlCeConnection(FormMain.SqlConStr);
            con.Open();
            SqlCeCommand cmd = con.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM tblEmail WHERE Address = '" + address + "'";
            int result = ((int)cmd.ExecuteScalar());
            con.Close();
            con.Dispose();
            if (result == 0)
                return false;
            return true;
        }

        private List<MailAddress> GetRecipients(string type)
        {
            try
            {
                List<MailAddress> recip = new List<MailAddress>() { };

                if (type != null)
                {
                    if (!String.IsNullOrEmpty(type))
                    {
                        if (type == "Quick")
                        {
                            for (int i = 0; i < emailDb.Rows.Count; i++)
                                if (Convert.ToBoolean(emailDb.Rows[i]["Quick"]))
                                    recip.Add(new MailAddress(emailDb.Rows[i]["Address"].ToString(), emailDb.Rows[i]["Name"].ToString()));
                        }
                        else
                        {
                            for (int i = 0; i < emailDb.Rows.Count; i++)
                                if (emailDb.Rows[i]["Type"].ToString() == type)
                                    recip.Add(new MailAddress(emailDb.Rows[i]["Address"].ToString(), emailDb.Rows[i]["Name"].ToString()));
                        }
                    }
                    else
                        for (int i = 0; i < emailDb.Rows.Count; i++)
                            recip.Add(new MailAddress(emailDb.Rows[i]["Address"].ToString(), emailDb.Rows[i]["Name"].ToString()));
                }

                return recip;
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
            return new List<MailAddress>() { };
        }

        public bool Send(string attachment, DateTime date, string type, string emailHeader, string emailBody)
        {
            try
            {
                Log.n("Forbereder sending av epost.. (" + type + ")");

                List<MailAddress> recip = GetRecipients(type);
                if (recip.Count == 0)
                {
                    Log.n("Fant ingen mottakere for denne gruppen (" + type + ")");
                    return false;
                }

                string tittel = FilterString(emailHeader, date);
                string tekst = FilterString(emailBody, date);

                if (main.appConfig.epostNesteMelding.Length > 5 && type != "Quick")
                    tekst = main.appConfig.epostNesteMelding;

                if (InternalSendMail(recip, tittel, tekst, attachment, main.appConfig.epostBrukBcc))
                {
                    Log.n("E-post sendt for mottakere '" + type + "'.", Color.Green);
                    if (main.appConfig.epostNesteMelding.Length > 0 && type != "Quick")
                        main.appConfig.epostNesteMelding = ""; // slett melding
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil ved sending av epost", ex);
                errorMsg.ShowDialog();
                return false;
            }
        }

        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public bool InternalSendMail(List<MailAddress> recipient, string subject, string body, string attachmentFilename, bool useBcc = true)
        {
            try
            {
                string str = "";
                foreach (MailAddress mail in recipient)
                    str += mail.Address;

                string from = main.appConfig.epostAvsender;
                if (!IsValidEmail(from))
                    from = "no-reply@elkjop.no";

                Log.n("Sender epost til " + str, null, true);
                Log.n("Sender epost med vedlegg: " + attachmentFilename, null, true);
                if (useBcc)
                {
                    foreach (MailAddress mailaddress in recipient)
                    {
                        SmtpClient smtpClient = new SmtpClient();
                        MailMessage message = new MailMessage();
                        MailAddress fromAddress;
                        if (String.IsNullOrEmpty(main.appConfig.epostAvsenderNavn))
                            fromAddress = new MailAddress(main.appConfig.epostAvsender);
                        else
                            fromAddress = new MailAddress(main.appConfig.epostAvsender, main.appConfig.epostAvsenderNavn);

                        smtpClient.Host = main.appConfig.epostSMTPserver;
                        smtpClient.Port = main.appConfig.epostSMTPport;
                        smtpClient.EnableSsl = main.appConfig.epostSMTPssl;
                        smtpClient.Timeout = (30 * 1000);

                        message.From = fromAddress;
                        message.Subject = subject;
                        message.IsBodyHtml = true;
                        message.Body = body;
                        message.To.Add(mailaddress);

                        if (attachmentFilename != null)
                        {
                            Attachment attachment = new Attachment(attachmentFilename, MediaTypeNames.Application.Octet);
                            ContentDisposition disposition = attachment.ContentDisposition;
                            disposition.CreationDate = File.GetCreationTime(attachmentFilename);
                            disposition.ModificationDate = File.GetLastWriteTime(attachmentFilename);
                            disposition.ReadDate = File.GetLastAccessTime(attachmentFilename);
                            disposition.FileName = Path.GetFileName(attachmentFilename);
                            disposition.Size = new FileInfo(attachmentFilename).Length;
                            disposition.DispositionType = DispositionTypeNames.Attachment;
                            message.Attachments.Add(attachment);
                        }

                        smtpClient.Send(message);
                        Log.d("Fullført sending av e-post til " + mailaddress.Address);
                    }
                }
                else
                {
                    // Sendt alt ut som en, dette vil vise alle andre mottakere
                    SmtpClient smtpClient = new SmtpClient();
                    MailMessage message = new MailMessage();
                    MailAddress fromAddress;
                    if (String.IsNullOrEmpty(main.appConfig.epostAvsenderNavn))
                        fromAddress = new MailAddress(main.appConfig.epostAvsender);
                    else
                        fromAddress = new MailAddress(main.appConfig.epostAvsender, main.appConfig.epostAvsenderNavn);

                    smtpClient.Host = main.appConfig.epostSMTPserver;
                    smtpClient.Port = main.appConfig.epostSMTPport;
                    smtpClient.Timeout = (30 * 1000);

                    message.From = fromAddress;
                    message.Subject = subject;
                    message.IsBodyHtml = true;
                    message.Body = body;
                    foreach (MailAddress mailaddress in recipient)
                        message.To.Add(mailaddress);

                    if (attachmentFilename != null)
                    {
                        Attachment attachment = new Attachment(attachmentFilename, MediaTypeNames.Application.Octet);
                        ContentDisposition disposition = attachment.ContentDisposition;
                        disposition.CreationDate = File.GetCreationTime(attachmentFilename);
                        disposition.ModificationDate = File.GetLastWriteTime(attachmentFilename);
                        disposition.ReadDate = File.GetLastAccessTime(attachmentFilename);
                        disposition.FileName = Path.GetFileName(attachmentFilename);
                        disposition.Size = new FileInfo(attachmentFilename).Length;
                        disposition.DispositionType = DispositionTypeNames.Attachment;
                        message.Attachments.Add(attachment);
                    }

                    smtpClient.Send(message);

                    Log.d("Fullført sending av e-post.");
                }
                return true;
            }
            catch(SmtpException ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil oppstod under kommunikasjon med SMTP server " + main.appConfig.epostSMTPserver + ". Se logg for detaljer.");
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
            return false;
        }


    }
}
