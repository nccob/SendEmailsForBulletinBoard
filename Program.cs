using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Configuration;
using System.Data;
//using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Web;
using System.Data.SqlClient;
using Microsoft.Win32;

namespace SendEmailsFromNCCOB
{
    class Program
    {
        private static List<MessageMeta> MessageList;
        private static RegistryKey regkey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\NCCOBOnline");

        private static string _dbServer = System.Configuration.ConfigurationManager.AppSettings["server"];
        private static string _dbName = System.Configuration.ConfigurationManager.AppSettings["database"];
        private static string _dbUser = regkey.GetValue("dbUser").ToString();
        private static string _dbPassword = regkey.GetValue("dbPassword").ToString();

        private const int COMMAND_TIMEOUT = 600;
        private const int CONNECTION_TIMEOUT = 120;
        private static string _dbConnString;



        static void Main(string[] args)
        {
            //TODO: Add error handling
            
                ProcessUnsentEmails();
           
            
        }

        private static void ProcessUnsentEmails()
        { 
            //get dataset of emails to send
            GetData();
            
            //For each MessageMeta object call Send Email
            foreach (MessageMeta _mm in MessageList)
            {
                SendEmail(_mm);
            }
        }

        private static void GetData()
        {
            ////OdbcConnection DbConnection = null;
            ////OdbcConnection DbConnection2 = null;
            ////OdbcCommand DbCommand = null;
            ////OdbcDataReader DbReader = null;
            ////OdbcCommand DbCommand2 = null;
            ////OdbcDataReader DbReader2 = null;

            SqlConnection DbConnection = null;
            SqlConnection DbConnection2 = null;
            SqlCommand DbCommand = null;
            SqlDataReader DbReader = null;
            SqlCommand DbCommand2 = null;
            SqlDataReader DbReader2 = null;

            try
            {

                 _dbConnString = "server=\'" + _dbServer + "\'; user id=\'" +
                    _dbUser + "\'; password=\'" + _dbPassword + "\'; Database=\'" +
                    _dbName + "\';connection timeout=" + CONNECTION_TIMEOUT +
                    "; MultipleActiveResultSets=True; Max Pool Size = 1000; Pooling = True;";


                //DbConnection = new OdbcConnection("DSN=" + ConfigurationManager.AppSettings["NCCOB_DSN"].ToString() + ";UID=" + UID + ";PWD=" + PWD + ";");
                //DbConnection.Open();
                DbConnection = new SqlConnection(_dbConnString);
                DbConnection.Open();

                DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = "SELECT * FROM EmailsFromApp where SentDate is null and ErrorCount <= 5 and ToEmail not like '%@test%' and ToEmail not like '%.test.%' "; 
                DbReader = DbCommand.ExecuteReader();

                //Create list of MessageMeta object
                MessageList = new List<MessageMeta>();
                while (DbReader.Read())
                {
                    MessageMeta _mm = new MessageMeta();
                    _mm.ID = Convert.ToInt32(DbReader["ID"]);
                    _mm.BCC = DbReader["BCC"] != null ? DbReader["BCC"].ToString() : "";
                    _mm.CC = DbReader["CC"] != null ? DbReader["CC"].ToString() : "";
                    _mm.EmailText = DbReader["EmailText"] != null ? HttpUtility.HtmlDecode(DbReader["EmailText"].ToString()) : "";
                    _mm.From = DbReader["FromEmail"] != null ? DbReader["FromEmail"].ToString() : "";
                    _mm.Subject = DbReader["Subject"] != null ? DbReader["Subject"].ToString() : "";
                    _mm.To = DbReader["ToEmail"] != null ? DbReader["ToEmail"].ToString() : "";
                    _mm.ErrorCount = DbReader.IsDBNull(DbReader.GetOrdinal("ErrorCount")) ? 0 : Convert.ToInt32(DbReader["ErrorCount"]);

                    //clean up the To when there are multiple email addresses separated by semicolons.
                    _mm.To = _mm.To.Replace(';', ',');

                    //DbConnection2 = new OdbcConnection("DSN=" + ConfigurationManager.AppSettings["NCCOB_DSN"].ToString() + ";UID=" + UID + ";PWD=" + PWD + ";");
                    DbConnection2 = new SqlConnection(_dbConnString);
                    DbConnection2.Open();
                    DbCommand2 = DbConnection2.CreateCommand();
                    DbCommand2.CommandText = "SELECT * FROM EmailsFromAppAttachment where EmailsFromAppID = " + _mm.ID.ToString();
                    DbReader2 = DbCommand2.ExecuteReader();
                    while (DbReader2.Read())
                    {
                        string _fileName = DbReader2["FileName"].ToString();
                        Byte[] _fileData  = (Byte[])DbReader2["FileData"];
                        MessageAttachment _a = new MessageAttachment();
                        _a.FileName = _fileName;
                        _a.FileData = _fileData;
                        _mm.FilesToAttach.Add(_a);
                    }

                    MessageList.Add(_mm);
                }
            }
            catch (Exception ex)
            {
                string err = "Error Caught in SendEmailsFromNCCOB application\n" +
                        "Error in: GetData()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);
            }
            finally
            {
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();
                if (DbReader2 != null) DbReader2.Close();
                if (DbCommand2 != null) DbCommand2.Dispose();
                if (DbConnection2 != null) DbConnection2.Close();
            }
        }
               

        private static void SendEmail(MessageMeta _meta)
        {
            try
            {
                if (_meta.To.Replace(" ", "") == string.Empty)
                    throw new Exception("Missing 'To Email Address'");

                if (_meta.From != string.Empty && _meta.To != string.Empty)
                {
                    MailMessage message = new MailMessage(_meta.From, _meta.To, _meta.Subject, _meta.EmailText);
                    message.IsBodyHtml = true;
                    if (_meta.CC != string.Empty)
                        message.CC.Add(_meta.CC);

                    if (_meta.BCC != string.Empty)
                        message.Bcc.Add(_meta.BCC);

                    if (_meta.FilesToAttach != null && _meta.FilesToAttach.Count > 0)
                    {
                        foreach (MessageAttachment _a in _meta.FilesToAttach)
                        {
                            MemoryStream memoryStream = new MemoryStream();
                            memoryStream.Write(_a.FileData, 0, _a.FileData.Length);
                            memoryStream.Seek(0, SeekOrigin.Begin);

                            Attachment _attachment = new Attachment(memoryStream, _a.FileName);
                            message.Attachments.Add(_attachment);
                        }
                    }

                    SmtpClient emailClient;
                    emailClient = new SmtpClient(ConfigurationManager.AppSettings["smtpserver"], 25);

                    emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    emailClient.Send(message);

                    Update(_meta, false);

                }
            }
            catch (Exception ex)
            {
                string err = "Error Caught in SendEmailsFromNCCOB application\n" +
                        "Error in: SendEmail()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);

                _meta.ErrorCount++;

                Update(_meta, true);

                //send the message to the developers only on the fifth error.  That way the developers may not have to get involved.
                if (_meta.ErrorCount == 5)
                {
                    //Need to do a little adjusting and to notify IT
                    string note = @"This email failed when sending to <b>'" + _meta.To + @"'</b>.  Take a look to see if you can determine 
                                what's wrong with the email address and then manually send if you can.  The application will try to resend the original 
                                email for about 3 minutes.<br/><hr/><br/><b>Subject:</b> " + _meta.Subject;

                    _meta.To = "NCCOBDevelopers@nccob.gov";
                    //_meta.To = "ssnively@nccob.gov";
                    _meta.EmailText = note + "<br/><b>Email Body:</b> " + _meta.EmailText;
                    _meta.Subject = "Auto Email Send Error";

                    SendEmail(_meta);

                }
                                
            }
        
           
        }

        private static void Update(MessageMeta _meta, bool HadError)
        {
            SqlConnection DbConnection = null;
            SqlCommand DbCommand = null;
            DbConnection = new SqlConnection(_dbConnString);
            DbConnection.Open();

            //OdbcConnection DbConnection = null;
            //OdbcCommand DbCommand = null;
            //DbConnection = new OdbcConnection("DSN=" + ConfigurationManager.AppSettings["NCCOB_DSN"].ToString() + ";UID=" + UID + ";PWD=" + PWD + ";");
            //DbConnection.Open();
            //System.Data.Odbc.OdbcTransaction _tran = DbConnection.BeginTransaction();
                
            try
            {
                string sql = "";

                if (HadError)
                    sql = "Update EmailsFromApp set haderror = 1, ErrorCount= " + _meta.ErrorCount.ToString() + " where id = " + _meta.ID.ToString();
                else
                    sql = "Update EmailsFromApp set SentDate = getdate() where id = " + _meta.ID.ToString();


                DbCommand = DbConnection.CreateCommand();
                //DbCommand.Transaction = _tran;
                DbCommand.CommandText = sql;
                DbCommand.ExecuteNonQuery();

                //throw new Exception("test");

                //_tran.Commit();
                
            }
            catch (Exception ex)
            {
                //_tran.Rollback();
                string err = "Error Caught in SendEmailsFromNCCOB application\n" +
                        "Error in: Update()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("SendEmailsFromNCCOB", err, EventLogEntryType.Error);
            }
            finally
            {
                DbCommand.Dispose();
                DbConnection.Close();
            }
        }
    }

    class MessageMeta
    {
        public string EmailText { get; set; }
        public string Subject { get; set; } 
        public string To { get; set; }
        public string From { get; set; }
        public string CC { get; set; }
        public string BCC { get; set; }
        public int ID { get; set; }
        public List<MessageAttachment> FilesToAttach = new List<MessageAttachment>();

        public int ErrorCount { get; set; }
    }

    class MessageAttachment
    {
        public string FileName { get; set; }
        public Byte[] FileData { get; set; }
    }
}
