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

namespace SendEmailsForBulletinBoard
{
    class Program
    {
        private static RegistryKey regkey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\NCCOBOnline");

        private static string _dbServer = System.Configuration.ConfigurationManager.AppSettings["server"];
        private static string _dbName = System.Configuration.ConfigurationManager.AppSettings["database"];
        private static string _dbUser = regkey.GetValue("dbUser").ToString();
        private static string _dbPassword = regkey.GetValue("dbPassword").ToString();

        private const int COMMAND_TIMEOUT = 600;
        private const int CONNECTION_TIMEOUT = 120;
        private static string _dbConnString;

        private static int _RecipientsPerDayLimit;
        private static int _RecipientsPerEmailLimit;
        private static int _EmailsPerMinuteLimit;
        private static bool _AllowTestEmails = (_dbServer != "10.53.16.21");
        private static bool _ErrorRaised;

        private static int _RecipientsSentToday;
        private static int _EmailsForCompanyContactsWaiting;
        private static int _EmailsForSubscribersWaiting;
        private static int _EmailsToSendForCompanyContacts;
        private static int _EmailsToSendForSubscribers;

        private static SqlConnection DbConnection = null;
        private static SqlCommand DbCommand = null;
        private static SqlDataReader DbReader = null;

        static void Main(string[] args)
        {
            _dbConnString = "server=\'" + _dbServer + "\'; user id=\'" +
                    _dbUser + "\'; password=\'" + _dbPassword + "\'; Database=\'" +
                    _dbName + "\';connection timeout=" + CONNECTION_TIMEOUT +
                    "; MultipleActiveResultSets=True; Max Pool Size = 1000; Pooling = True;";

            ProcessUnsentEmails();
            
        }

        private static void ProcessUnsentEmails()
        {

            SetLimits();

            while (!_ErrorRaised)
            {
                GetCurrentTally();

                if (_RecipientsSentToday >= _RecipientsPerDayLimit)  //Daily Limit
                    break;
                
                if (_EmailsForCompanyContactsWaiting + _EmailsForSubscribersWaiting == 0) //No more to send
                    break;

                if (_EmailsForCompanyContactsWaiting > 0)
                    PrepareCompanyContactEmail();
                else
                    PrepareSubscriberEmail();

               // System.Threading.Thread.Sleep(10000); /*wait 10 seconds and run again so we never hit the email per minute limit*/
            }

        }

        private static void SetLimits()
        {  
            try
            {
                DbConnection = new SqlConnection(_dbConnString);
                DbConnection.Open();
                DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = "SELECT * FROM RefValues where RefType = 'EmailAccountLimits'";
                DbReader = DbCommand.ExecuteReader();
                while (DbReader.Read())
                {
                    switch (DbReader["RefName"].ToString())
                    {
                        case "RecipientsPerDay":
                            _RecipientsPerDayLimit = Convert.ToInt32(DbReader["RefValue"]);
                            break;
                        case "RecipientsPerEmail":
                            _RecipientsPerEmailLimit = Convert.ToInt32(DbReader["RefValue"]);
                            break;
                        case "EmailsPerMinute":
                            _EmailsPerMinuteLimit = Convert.ToInt32(DbReader["RefValue"]);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                _ErrorRaised = true;

                string err = "Error Caught in SendEmailsForBulletinBoard application\n" +
                        "Error in: SetLimits()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);

            }
            finally
            {
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();
            }
        }

        private static void GetCurrentTally()
        {
            try
            {
                DbConnection = new SqlConnection(_dbConnString);
                DbConnection.Open();
                DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = @"select count(*) [RecipientsSentToday] 
                                            from BBEmail
                                            where SentDate > convert(varchar(10), getdate(), 101)";

                DbReader = DbCommand.ExecuteReader();
                while (DbReader.Read())
                {
                    _RecipientsSentToday = Convert.ToInt32(DbReader["RecipientsSentToday"]);                           
                }
                if (DbReader != null) DbReader.Close();

                DbCommand.CommandText = @"select 
	                                        sum(case when ContactID is not null then 1 else 0 end) [EmailsToSendtoCompanyContacts],
	                                        sum(case when BBSubscriberID is not null then 1 else 0 end) [EmailsToSendtoSubscribers]
                                        from 
                                            BBEmail
                                        where 
                                            SentDate is null ";

                if (!_AllowTestEmails)
                {
                    DbCommand.CommandText += @"and EmailAddressTo not like '%@test%'
                                            and EmailAddressTo not like '%.test.%'";
                }

                DbReader = DbCommand.ExecuteReader();
                while (DbReader.Read())
                {
                    _EmailsForCompanyContactsWaiting = DbReader["EmailsToSendtoCompanyContacts"] != DBNull.Value ? Convert.ToInt32(DbReader["EmailsToSendtoCompanyContacts"]) : 0;
                    _EmailsForSubscribersWaiting = DbReader["EmailsToSendtoSubscribers"] != DBNull.Value ? Convert.ToInt32(DbReader["EmailsToSendtoSubscribers"]): 0;
                }

                _EmailsToSendForCompanyContacts = Math.Min(_RecipientsPerDayLimit - _RecipientsSentToday, _EmailsForCompanyContactsWaiting);
                _EmailsToSendForSubscribers = Math.Min(_RecipientsPerDayLimit - _RecipientsSentToday - _EmailsToSendForCompanyContacts, _EmailsForSubscribersWaiting);
            }
            catch (Exception ex)
            {
                _ErrorRaised = true;

                string err = "Error Caught in SendEmailsForBulletinBoard application\n" +
                        "Error in: GetCurrentTally()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);
            }
            finally
            {
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();
            }
        }



        private static void PrepareCompanyContactEmail()
        {  
            try
            {
                DbConnection = new SqlConnection(_dbConnString);
                
                string sql = "exec mlsadmin.BBRetrieveCompanyContactEmailList " + _EmailsToSendForCompanyContacts.ToString() + ";";

                DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = sql;
                //DbCommand.CommandType = CommandType.StoredProcedure;

                //DbCommand.Parameters.Add(new SqlParameter("CountToRetrieve", _EmailsToSendForCompanyContacts));

                DbConnection.Open();
                DbReader = DbCommand.ExecuteReader();

                while (DbReader.Read())
                {
                    MessageMeta _mm = new MessageMeta();
                    _mm.To = DbReader["EmailAddressFrom"] != DBNull.Value ? DbReader["EmailAddressFrom"].ToString() : "";  /*Yes, the from is the to since we are using the BCC*/
                    _mm.From = DbReader["EmailAddressFrom"] != DBNull.Value ? DbReader["EmailAddressFrom"].ToString() : "";
                    _mm.BCC = DbReader["BCCList"] != DBNull.Value ? DbReader["BCCList"].ToString() : "";
                    _mm.CC = System.Configuration.ConfigurationManager.AppSettings["ccemailaddress"];
                    _mm.EmailText = DbReader["Body"] != DBNull.Value ? HttpUtility.HtmlDecode(DbReader["Body"].ToString()) : "";
                    _mm.Subject = DbReader["EmailSubject"] != DBNull.Value ? DbReader["EmailSubject"].ToString() : "";

                    bool _sendEmailWorked;
                    _sendEmailWorked = SendEmail(_mm);

                    if (!_sendEmailWorked)
                        throw new Exception("There was an error sending the Company Contact Emails: " + _mm.BCC);
                }
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();

                DbConnection.Open();
                DbCommand.CommandText = @"update BBEmail set SentDate=getdate(), FlagToSend=0 where FlagToSend = 1 and ContactID is not null";
                DbCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                _ErrorRaised = true;

                string err = "Error Caught in SendEmailsForBulletinBoard application\n" +
                        "Error in: PrepareCompanyContactEmail()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);
            }
            finally
            {
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();
                
            }
        }

        private static void PrepareSubscriberEmail()
        {
            try
            {                
                string UnsubscribeURL = System.Configuration.ConfigurationManager.AppSettings["UnsubscribeURL"];

                DbConnection = new SqlConnection(_dbConnString);
                DbConnection.Open();

                DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = @"exec mlsadmin.BBRetrieveSubscriberEmailList " + _EmailsToSendForSubscribers.ToString() + ",'" + UnsubscribeURL + "';";
                DbReader = DbCommand.ExecuteReader();
                                
                while (DbReader.Read())
                {                   
                    MessageMeta _mm = new MessageMeta();
                    _mm.To = DbReader["EmailAddressTo"] != DBNull.Value ? DbReader["EmailAddressTo"].ToString() : "";
                    _mm.From = DbReader["EmailAddressFrom"] != DBNull.Value ? DbReader["EmailAddressFrom"].ToString() : "";
                    //_mm.BCC = System.Configuration.ConfigurationManager.AppSettings["ccemailaddress"];
                    _mm.EmailText = DbReader["Body"] != DBNull.Value ? HttpUtility.HtmlDecode(DbReader["Body"].ToString()) : "";
                    _mm.Subject = DbReader["EmailSubject"] != DBNull.Value ? DbReader["EmailSubject"].ToString() : "";

                    string _sql;
                    bool _sendEmailWorked;
                    _sendEmailWorked = SendEmail(_mm);

                    if (_sendEmailWorked)
                        _sql = "update BBEmail set SentDate = getdate(), FlagToSend = 0 where FlagToSend = 1 and id = " + DbReader["ID"].ToString();
                    else
                        _sql = "update BBEmail set FlagToSend = 0 where FlagToSend = 1 and id = " + DbReader["ID"].ToString();


                    SqlCommand DbCommand2 = null;
                    DbCommand2 = DbConnection.CreateCommand();
                    DbCommand2.CommandText = _sql;
                    DbCommand2.ExecuteNonQuery();
                    DbCommand2.Dispose();
                }


            }
            catch (Exception ex)
            {
                _ErrorRaised = true;

                string err = "Error Caught in SendEmailsForBulletinBoard application\n" +
                        "Error in: PrepareSubscriberEmail()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString();
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);
            }
            finally
            {
                if (DbReader != null) DbReader.Close();
                if (DbCommand != null) DbCommand.Dispose();
                if (DbConnection != null) DbConnection.Close();

            }
        }



        private static bool SendEmail(MessageMeta _meta)
        {
            try
            {
                if (_meta.To.Replace(" ", "") == string.Empty)
                    throw new Exception("Missing 'To Email Address'");

                if (_meta.From != string.Empty && _meta.To != string.Empty)
                {
                    MailMessage message = new MailMessage(_meta.From, _meta.To, _meta.Subject, _meta.EmailText);
                    message.IsBodyHtml = true;
                    if (!String.IsNullOrWhiteSpace(_meta.CC))
                        message.CC.Add(_meta.CC);

                    if (!String.IsNullOrWhiteSpace(_meta.BCC))
                        message.Bcc.Add(_meta.BCC);                    

                    SmtpClient emailClient;
                    emailClient = new SmtpClient(ConfigurationManager.AppSettings["smtpserver"], 25);

                    emailClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    emailClient.Send(message);

                    int _delaymiliseconds = 60/_EmailsPerMinuteLimit * 1000;
                    System.Threading.Thread.Sleep(_delaymiliseconds); /*wait so we never hit the email per minute limit*/

                }

                return true;
            }
            catch (Exception ex)
            {
                _ErrorRaised = true;

                string err = "Error Caught in SendEmailsForBulletinBoard application\n" +
                        "Error in: SendEmail()" +
                        "\nError Message:" + ex.Message.ToString() +
                        "\nStack Trace:" + ex.StackTrace.ToString() +
                        "\nEmailAddressTo:" + _meta.To;
                EventLog.WriteEntry("NCCOBApp", err, EventLogEntryType.Error);

                return false;

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
        
    }


}
