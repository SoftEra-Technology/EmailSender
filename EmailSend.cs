using ProcsDLL.Models.BoardMeeting.Model;
using ProcsDLL.Models.BoardMeeting.Repository;
using ProcsDLL.Models.BoardMeeting.Service.Request;
using ProcsDLL.Models.BoardMeeting.Service.Response;
using ProcsDLL.Models.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;

namespace EmailSender
{
    public class EmailSend
    {
        static Meeting objMeeting;
        static Email objEmail;
        static MeetingRepository meetingRepo;
        static MeetingResponse meetingResponse;
        static CommitteeResponse committeeResponse;
        static UserResponse userResponse;
        static EmailResponse emailResponse;
        static EmailRequest emailRequest;
        static EmailRepository emailRepo;

        public static void Main(string[] args)
        {
            ////////PROCS_BOARD_MEETING 

            ////SendMeetingNoticeEmail(Int32 companyId, Int32 meetingId, String moduleDatabase, String subject, String emailToList, String fileName, String layoutTemplate, String templateMode)

            //CreateDraftMOMTaskAndSendEmail(Int32 companyId, Int32 committeeId, Int32 meetingId, Int32 draftMOMId, String userLogin, String dataBase, Int32 userId, String mode)
            //CreateDraftMOMTaskAndSendEmail(8, 44, 3022, 2, "ssingh", "PROCS_BOARD_MEETING", 1050, "NewDraft");


            //args = new String[10];
            //args[0] = "SENDMEETINGNOTICE";
            //args[1] = Convert.ToString(8);
            //args[2] = Convert.ToString(3022);
            //args[3] = Convert.ToString("PROCS_BOARD_MEETING");
            //args[4] = Convert.ToString("Meeting Scheduled");
            //args[5] = "romilnehru@outlook.com";
            //args[6] = "ABC";
            //args[7] = Convert.ToString("<p>Respected Sir/Madam2,</p><p>This is to inform you that 1st Meeting of Audit Committee of Audit Committee has been scheduled.</p><p>The venue of the meeting is at Delhi,Board Room, 3rd floor, ACME Corp Ltd, DLF Promnegate, Delhi on Jul 17, 2020 from 08:15 AM to 10:15 AM.</p><p>Thanks &amp; Regards,</p><p>Abhinav</p>");
            //args[8] = Convert.ToString("WithCalender");

            String mode = args[0];
            Int32 companyId = Convert.ToInt32(args[1]);

            try
            {
                if (args.Length > 0)
                {
                    switch (mode.ToUpper())
                    {
                        case "MEETINGRESCHEDULE":

                            SendMeetingRescheduleEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToString(args[3]));

                            break;

                        case "SUBMITITEMS":

                            SendSubmitItemsEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToString(args[3]));

                            break;

                        case "SENDMEETINGNOTICE":
                            //SendMeetingNoticeEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToString(args[3]), Convert.ToString(args[4]), Convert.ToString(args[5]), Convert.ToString(args[6]), Convert.ToString(args[7]), Convert.ToString(args[8]));
                            SendMeetingNoticeEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToString(args[3]), Convert.ToString(args[4]), Convert.ToString(args[5]), Convert.ToString(args[6]));

                            break;

                        case "SENDFINALMOMEMAILALERT":

                            SendFinalMOMEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToString(args[3]), Convert.ToString(args[4]));

                            break;

                        case "SENDATRTASKEMAIL":

                            SendATRTaskEmail(Convert.ToInt32(args[1]), Convert.ToString(args[2]), Convert.ToString(args[3]), Convert.ToString(args[4]));

                            break;

                        case "SENDCIRCULARRESOLUTIONEMAIL":

                            SendCREmail(Convert.ToInt32(args[1]), Convert.ToString(args[2]), Convert.ToString(args[3]), Convert.ToInt32(args[4]), Convert.ToString(args[5]), Convert.ToInt32(args[6]), Convert.ToString(args[7]), Convert.ToString(args[8]), Convert.ToString(args[9]));

                            break;

                        case "SENDCOMMENTEMAIL":

                            SendCommentsShareEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]), Convert.ToInt32(args[4]), Convert.ToString(args[5]), Convert.ToString(args[6]), Convert.ToString(args[7]));

                            break;

                        case "SENDDRAFTMOMEMAIL":
                                                       
                            CreateDraftMOMTaskAndSendEmail(Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]), Convert.ToInt32(args[4]), Convert.ToString(args[5]), Convert.ToString(args[6]), Convert.ToInt32(args[7]), Convert.ToString(args[8]));

                            break;

                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", mode, "Admin", 1, companyId);
            }
        }

        #region "Send Meeting Reschedule Email"

        private static void SendMeetingRescheduleEmail(Int32 companyId, Int32 meetingId, String moduleDatabase)
        {
            try
            {
                objMeeting = new Meeting();
                objMeeting.ID = meetingId;
                objMeeting.moduleDatabase = moduleDatabase;
                objMeeting.companyId = companyId;

                objEmail = new Email();
                //objEmail.templateId = 3;
                objEmail.taskName = "Meeting Rescheduled";
                objEmail.isMeeting = true;
                objEmail.meeting = new Meeting();
                objEmail.meeting.ID = objMeeting.ID;
                objEmail.moduleDatabase = moduleDatabase;
                objEmail.companyId = companyId;
                objEmail.createdBy = "EmailSender Exe";

                meetingRepo = new MeetingRepository();
                meetingRepo.GetMeetingMembers(objMeeting);

                if (objMeeting.meetingMembers != null)
                {
                    if (objMeeting.meetingMembers.Count > 0)
                    {
                        objEmail.emailToList = new List<String>();
                        foreach (MeetingMembers member in objMeeting.meetingMembers)
                        {
                            objEmail.emailToList.Add(member.user.emailId);
                        }
                    }
                }

                emailResponse = new EmailResponse();
                emailRequest = new EmailRequest(objEmail);
                emailResponse = emailRequest.GetTemplateHeaderDetailByNameAndCompanyId();
                //emailResponse = emailRequest.GetTemplateHeaderDetailByTemplateId();
                if (emailResponse.StatusFl)
                {
                    if (emailResponse.EmailList != null)
                    {
                        if (emailResponse.EmailList.Count > 0)
                        {
                            objEmail.layoutTemplate = emailResponse.EmailList[0].layoutTemplate;
                            objEmail.subject = emailResponse.EmailList[0].subject;
                            objEmail.meeting = new Meeting();
                            objEmail.meeting = objMeeting;
                            //emailRequest = new EmailRequest(objEmail);
                            //emailResponse = emailRequest.SendMeetingRescheduleAlert();
                            EmailHelper.SendRescheduleMeetingEmail(objEmail);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendMeetingRescheduleEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send Submit Items Email"

        private static void SendSubmitItemsEmail(Int32 companyId, Int32 meetingId, String moduleDatabase)
        {
            try
            {
                objMeeting = new Meeting();
                objMeeting.ID = meetingId;
                objMeeting.moduleDatabase = moduleDatabase;

                objEmail = new Email();
                //objEmail.templateId = 4;
                objEmail.taskName = "Items Submitted";
                objEmail.isMeeting = true;
                objEmail.meeting = new Meeting();
                objEmail.meeting.ID = objMeeting.ID;
                objEmail.moduleDatabase = moduleDatabase;
                objEmail.companyId = companyId;

                userResponse = new UserResponse();
                meetingRepo = new MeetingRepository();
                userResponse = meetingRepo.GetCommitteeSuperAdmin(objMeeting);

                if (userResponse.StatusFl)
                {
                    objEmail.emailToList = new List<String>();
                    if (userResponse.User != null)
                    {
                        if (!String.IsNullOrEmpty(Convert.ToString(userResponse.User.emailId)))
                        {
                            objEmail.emailToList.Add(userResponse.User.emailId);
                        }
                    }

                    emailResponse = new EmailResponse();
                    emailRequest = new EmailRequest(objEmail);
                    emailResponse = emailRequest.GetTemplateHeaderDetailByNameAndCompanyId();
                    //emailResponse = emailRequest.GetTemplateHeaderDetailByTemplateId();
                    if (emailResponse.StatusFl)
                    {
                        if (emailResponse.EmailList != null)
                        {
                            if (emailResponse.EmailList.Count > 0)
                            {
                                objEmail.layoutTemplate = emailResponse.EmailList[0].layoutTemplate;
                                objEmail.subject = emailResponse.EmailList[0].subject;
                                objEmail.meeting = new Meeting();
                                objEmail.meeting = objMeeting;
                                emailRequest = new EmailRequest(objEmail);
                                emailResponse = emailRequest.SendMeetingRescheduleAlert();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendSubmitItemsEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send Meeting Notice Email"

        //private static void SendMeetingNoticeEmail(Int32 companyId, Int32 meetingId, String moduleDatabase, String subject, String emailToList, String fileName, String layoutTemplate, String templateMode)
        private static void SendMeetingNoticeEmail(Int32 companyId, Int32 meetingId, String moduleDatabase, String emailSenderParameterId, String fileName, String templateMode)
        {
            MeetingRepository meetingRepo = new MeetingRepository();
            try
            {
                objEmail = new Email();
                objEmail.templateMode = templateMode;
                objEmail.createdBy = "EmailSender Exe";
                objEmail.companyId = companyId;
                objEmail.meeting = new Meeting
                {
                    ID = meetingId,
                    companyId = companyId
                };
                objEmail.meeting = meetingRepo.GetMeetingForMeetingNotice(objEmail.meeting);
                objEmail.moduleDatabase = moduleDatabase;
                objEmail.emailSenderParameterId = Convert.ToInt32(emailSenderParameterId);

                if (!String.IsNullOrEmpty(Convert.ToString(fileName)))
                {
                    if (Convert.ToString(fileName) != "ABC")
                    {
                        objEmail.listAttachment = new List<String>();
                        String applicationPath = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["EmailFilePath"].ToString(), true);
                        String filePath = applicationPath + Convert.ToString(fileName);
                        objEmail.listAttachment.Add(filePath);
                        objEmail.attachment = Convert.ToString(fileName);
                    }
                    else
                    {
                        objEmail.listAttachment = new List<String>();
                        objEmail.attachment = String.Empty;
                        //new LogHelper().AddExceptionLogs("EmailSender", "1", "no file", "", "", "", 1, companyId);
                    }
                }

                //objEmail.subject = subject;
                //if (!String.IsNullOrEmpty(emailToList))
                //{
                //    objEmail.emailToList = new List<String>();
                //    objEmail.emailToList = Convert.ToString(emailToList).Split(',').ToList();
                //}
                //objEmail.layoutTemplate = layoutTemplate;

                emailResponse = new EmailResponse();
                emailRequest = new EmailRequest(objEmail);
                emailResponse = emailRequest.GetEmailSenderParameters();
                if (emailResponse.StatusFl)
                {                    
                    objEmail = new Email();
                    objEmail = emailResponse.Email;
                    if (!String.IsNullOrEmpty(objEmail.emailTo))
                    {
                        objEmail.emailToList = new List<String>();
                        objEmail.emailToList = Convert.ToString(objEmail.emailTo).Split(',').ToList();
                    }
                    //new LogHelper().AddExceptionLogs("1", emailResponse.Email.emailTo, "", emailResponse.StatusFl.ToString(), "", "", 1, companyId);
                    //new LogHelper().AddExceptionLogs("2", emailResponse.Email.layoutTemplate, "", emailResponse.StatusFl.ToString(), "", "", 1, companyId);
                    //new LogHelper().AddExceptionLogs("3", emailResponse.Email.attachment, "", emailResponse.StatusFl.ToString(), "", "", 1, companyId);
                    //new LogHelper().AddExceptionLogs("4", emailResponse.Email.subject, "", emailResponse.StatusFl.ToString(), "", "", 1, companyId);
                    //new LogHelper().AddExceptionLogs("5", Convert.ToString(emailResponse.Email.meeting.ID), "", emailResponse.StatusFl.ToString(), "", "", 1, companyId);                  

                    emailResponse = new EmailResponse();
                    emailRequest = new EmailRequest(objEmail);
                    emailResponse = emailRequest.AddMeetingNoticeActivityLog();
                    //new LogHelper().AddExceptionLogs("EmailSender", "1", "status", emailResponse.StatusFl.ToString(), "", "", 1, companyId);
                    if (emailResponse.StatusFl)
                    {
                        //emailRequest = new EmailRequest(objEmail);
                        //emailResponse = emailRequest.SendEmailNotification();
                        EmailHelper.SendMeetingNoticeEmail(objEmail);
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendMeetingNoticeEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send Final MOM Email"

        private static void SendFinalMOMEmail(Int32 companyId, Int32 meetingId, String moduleDatabase, String userLogin)
        {
            try
            {
                objMeeting = new Meeting();
                objMeeting.ID = meetingId;
                objMeeting.companyId = companyId;
                objMeeting.moduleDatabase = moduleDatabase;
                objMeeting.createdBy = userLogin;

                meetingResponse = new MeetingResponse();
                meetingRepo = new MeetingRepository();
                meetingResponse = meetingRepo.listmeetingFinalMOM_upload(objMeeting);
                if (meetingResponse.StatusFl)
                {
                    MeetingFinalMOM objFinalMOM = meetingResponse.Meeting.meetingFinalMOMItems;

                    bool sentEmailInBulk = false;

                    IndexConfigurationRepository indexConfigRepo = new IndexConfigurationRepository();
                    IndexConfigurationResponse indexConfigResponse = new IndexConfigurationResponse();
                    PageConfig objPageConfig = new PageConfig();
                    objPageConfig.companyId = companyId;
                    objPageConfig.moduleDatabase = moduleDatabase;
                    indexConfigResponse = indexConfigRepo.GetPageWiseNumberConfig(objPageConfig);
                    if (indexConfigResponse.StatusFl)
                    {
                        if (indexConfigResponse.PageConfigList != null)
                        {
                            if (indexConfigResponse.PageConfigList.Count > 0)
                            {
                                if (indexConfigResponse.PageConfigList[0].isSendFinalMOMEmailinBulk == "Yes")
                                {
                                    sentEmailInBulk = true;
                                }
                            }
                        }
                    }

                    objEmail = new Email();
                    objEmail.companyId = companyId;
                    objEmail.createdBy = userLogin;
                    objEmail.meeting = new Meeting
                    {
                        ID = objFinalMOM.id
                    };
                    objEmail.moduleDatabase = moduleDatabase;

                    if (objFinalMOM.listSharedUsers != null)
                    {
                        if (objFinalMOM.listSharedUsers.Count > 0)
                        {
                            objEmail.emailToList = new List<String>();
                            if (sentEmailInBulk)
                            {
                                objEmail.emailToList.Add(objFinalMOM.listSharedUsers.Select(x => x.emailId).Aggregate((x, y) => x + "," + y));
                            }
                            else
                            {
                                foreach (User objUser in objFinalMOM.listSharedUsers)
                                {
                                    objEmail.emailToList.Add(objUser.emailId);
                                }
                            }
                        }
                    }

                    if (!String.IsNullOrEmpty(Convert.ToString(objFinalMOM.finalMOMFile)))
                    {
                        String applicationPath = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FinalMOMFilePath"].ToString(), true);
                        String filePath = applicationPath + objFinalMOM.finalMOMFile;
                        objEmail.listAttachment = new List<String>();
                        objEmail.listAttachment.Add(filePath);
                    }

                    //objEmail.templateId = 5;
                    objEmail.taskName = "Final MOM Uploaded";
                    objEmail.isMeeting = true;

                    emailResponse = new EmailResponse();
                    emailRequest = new EmailRequest(objEmail);
                    emailResponse = emailRequest.GetTemplateHeaderDetailByNameAndCompanyId();
                    //emailResponse = emailRequest.GetTemplateHeaderDetailByTemplateId();
                    if (emailResponse.StatusFl)
                    {
                        if (emailResponse.EmailList != null)
                        {
                            if (emailResponse.EmailList.Count > 0)
                            {
                                objEmail.layoutTemplate = emailResponse.EmailList[0].layoutTemplate;
                                objEmail.subject = emailResponse.EmailList[0].subject;
                                //new LogHelper().AddExceptionLogs("EmailSender", "1", "", "", "", "", 1);
                                //objEmail.meeting = new Meeting();
                                //objEmail.meeting = objMeeting;
                                //emailRepo = new EmailRepository();
                                //emailResponse = emailRepo.SendEmail(objEmail);
                                EmailHelper.SendFinalMOMEmail(objEmail);
                                //new LogHelper().AddExceptionLogs("EmailSender", "2", "", "", "", "", 1);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendFinalMOMEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send ATR Task Email"

        private static void SendATRTaskEmail(Int32 companyId, String userLogin, String moduleDatabase, String atrIds)
        {
            try
            {
                String defaultEmail = String.Empty, disclosureEmail = String.Empty, smtpHostName = String.Empty, smtpUserName = String.Empty, password = String.Empty;
                Int32 port = 0;
                bool ssl = false, userDefaultCredential = false, isSMTPConfigExists = false;

                ATRTaskResponse atrTaskRes = new ATRTaskResponse();
                ATRTaskRepository atrTaskRepo = new ATRTaskRepository();
                atrTaskRes = atrTaskRepo.GetATRAssignedToUsers(atrIds, userLogin, moduleDatabase, companyId);

                objEmail = new Email();
                objEmail.companyId = companyId;
                objEmail.moduleDatabase = moduleDatabase;
                if (atrTaskRes.StatusFl)
                {
                    if (atrTaskRes.ObservationTaskList != null)
                    {
                        if (atrTaskRes.ObservationTaskList.Count > 0)
                        {
                            using (SqlConnection conn = new SqlConnection(CryptorEngine.Decrypt(ConfigurationManager.AppSettings["ConnectionString"].ToString(), true)))
                            {
                                conn.Open();
                                conn.ChangeDatabase(moduleDatabase);
                                using (SqlCommand cmd = new SqlCommand("SP_PROCS_INSIDER_CONFIG_SMTP", conn))
                                {
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.CommandTimeout = 0;
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.Add(new SqlParameter("@Mode", "GET_Smtp_Config_List"));
                                    cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add(new SqlParameter("@COMPANY_ID", companyId));
                                    SqlDataReader rdr = cmd.ExecuteReader();
                                    if (rdr.HasRows)
                                    {
                                        while (rdr.Read())
                                        {
                                            isSMTPConfigExists = true;
                                            defaultEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["DEFAULT_EMAIL"]))) ? Convert.ToString(rdr["DEFAULT_EMAIL"]) : String.Empty;
                                            disclosureEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]))) ? Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]) : String.Empty;
                                            smtpHostName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_HOST_NAME"]))) ? Convert.ToString(rdr["SMTP_HOST_NAME"]) : String.Empty;
                                            port = Convert.ToInt32(rdr["PORT"]);
                                            ssl = (!String.IsNullOrEmpty(Convert.ToString(rdr["SSL"]))) ? (Convert.ToString(rdr["SSL"]) == "Yes" ? true : false) : false;
                                            smtpUserName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_USER_NAME"]))) ? Convert.ToString(rdr["SMTP_USER_NAME"]) : String.Empty;
                                            password = (!String.IsNullOrEmpty(Convert.ToString(rdr["PASSWORD"]))) ? CryptorEngine.Decrypt(Convert.ToString(rdr["PASSWORD"]), true) : String.Empty;
                                            userDefaultCredential = (!String.IsNullOrEmpty(Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]))) ? (Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]) == "Yes" ? true : false) : false;
                                        }
                                    }
                                    rdr.Close();
                                }
                                conn.Close();
                            }

                            if (isSMTPConfigExists)
                            {
                                foreach (ATREmail objATREmail in atrTaskRes.ObservationTaskList)
                                {
                                    EmailHelper.SendEmailToInitiateObservation(defaultEmail, objATREmail.assignedToEmail, smtpHostName, ssl, smtpUserName, password, userDefaultCredential, port, CryptorEngine.Encrypt(Convert.ToString(objATREmail.atrId), true), objATREmail, companyId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendATRTaskEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send Circular Resolution Email"

        private static void SendCREmail(Int32 companyId, String userLogin, String moduleDatabase, Int32 committeeId, String committeeName, Int32 crId, String crNo, String crDate, String crFile)
        {
            try
            {
                String defaultEmail = String.Empty, disclosureEmail = String.Empty, smtpHostName = String.Empty, smtpUserName = String.Empty, password = String.Empty;
                Int32 port = 0;
                bool ssl = false, userDefaultCredential = false, isSMTPConfigExists = false;

                crFile = CryptorEngine.Decrypt(Convert.ToString(ConfigurationManager.AppSettings["CircularResolutionFilePath"]), true) + crFile;

                CommitteeResponse committeeRes = new CommitteeResponse();
                CRRepository crRepo = new CRRepository();
                committeeRes = crRepo.GetCircularResolutionMembers(committeeId, companyId, moduleDatabase, crId);
                if (committeeRes.Committee != null)
                {
                    if (committeeRes.Committee.committeeMembers != null)
                    {
                        if (committeeRes.Committee.committeeMembers.Count > 0)
                        {
                            using (SqlConnection conn = new SqlConnection(CryptorEngine.Decrypt(ConfigurationManager.AppSettings["ConnectionString"].ToString(), true)))
                            {
                                conn.Open();
                                conn.ChangeDatabase(moduleDatabase);
                                using (SqlCommand cmd = new SqlCommand("SP_PROCS_INSIDER_CONFIG_SMTP", conn))
                                {
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.CommandTimeout = 0;
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.Add(new SqlParameter("@Mode", "GET_Smtp_Config_List"));
                                    cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add(new SqlParameter("@COMPANY_ID", companyId));
                                    SqlDataReader rdr = cmd.ExecuteReader();
                                    if (rdr.HasRows)
                                    {
                                        while (rdr.Read())
                                        {
                                            isSMTPConfigExists = true;
                                            defaultEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["DEFAULT_EMAIL"]))) ? Convert.ToString(rdr["DEFAULT_EMAIL"]) : String.Empty;
                                            disclosureEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]))) ? Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]) : String.Empty;
                                            smtpHostName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_HOST_NAME"]))) ? Convert.ToString(rdr["SMTP_HOST_NAME"]) : String.Empty;
                                            port = Convert.ToInt32(rdr["PORT"]);
                                            ssl = (!String.IsNullOrEmpty(Convert.ToString(rdr["SSL"]))) ? (Convert.ToString(rdr["SSL"]) == "Yes" ? true : false) : false;
                                            smtpUserName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_USER_NAME"]))) ? Convert.ToString(rdr["SMTP_USER_NAME"]) : String.Empty;
                                            password = (!String.IsNullOrEmpty(Convert.ToString(rdr["PASSWORD"]))) ? CryptorEngine.Decrypt(Convert.ToString(rdr["PASSWORD"]), true) : String.Empty;
                                            userDefaultCredential = (!String.IsNullOrEmpty(Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]))) ? (Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]) == "Yes" ? true : false) : false;
                                        }
                                    }
                                    rdr.Close();
                                }
                                conn.Close();
                            }

                            if (isSMTPConfigExists)
                            {
                                foreach (CommitteeMember objMember in committeeRes.Committee.committeeMembers)
                                {
                                    EmailHelper.SendEmailToInitiateCircularResolution(defaultEmail, objMember.member.emailId, smtpHostName, ssl, smtpUserName, password, userDefaultCredential, port, CryptorEngine.Encrypt(Convert.ToString(objMember.member.ID), true), committeeName, crNo, crDate, crFile, objMember.member.userName, companyId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendCREmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Send Comments Share Email"

        private static void SendCommentsShareEmail(Int32 companyId, Int32 committeeId, Int32 meetingId, Int32 agendaId, String commentIds, String userLogin, String moduleDatabase)
        {
            try
            {
                String defaultEmail = String.Empty, disclosureEmail = String.Empty, smtpHostName = String.Empty, smtpUserName = String.Empty, password = String.Empty;
                Int32 port = 0;
                bool ssl = false, userDefaultCredential = false, isSMTPConfigExists = false;

                //new LogHelper().AddExceptionLogs("EmailSender", "1", Convert.ToString(companyId), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "2", Convert.ToString(committeeId), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "3", Convert.ToString(meetingId), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "4", Convert.ToString(agendaId), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "5", Convert.ToString(commentIds), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "6", Convert.ToString(userLogin), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "7", Convert.ToString(moduleDatabase), "", "", "", 1, companyId);
                //new LogHelper().AddExceptionLogs("EmailSender", "7", "before call", "", "", "", 1, companyId);

                APICommentsSharing objComments = new APICommentsSharing();
                objComments.companyId = companyId;
                objComments.committeeId = committeeId;
                objComments.meetingId = meetingId;
                objComments.agendaId = agendaId;
                objComments.LoginId = userLogin;
                //new LogHelper().AddExceptionLogs("EmailSender", "8", "before call", "", "", "", 1, companyId);
                committeeResponse = new CommitteeResponse();
                meetingRepo = new MeetingRepository();
                //new LogHelper().AddExceptionLogs("EmailSender", "9", "before call", "", "", "", 1, companyId);

                committeeResponse = meetingRepo.GetCommentsSharedUsers(objComments, commentIds, moduleDatabase);

                objEmail = new Email();
                objEmail.companyId = companyId;
                objEmail.moduleDatabase = moduleDatabase;
                if (committeeResponse.StatusFl)
                {
                    if (committeeResponse.APICommitteeList != null)
                    {
                        if (committeeResponse.APICommitteeList.Count > 0)
                        {
                            using (SqlConnection conn = new SqlConnection(CryptorEngine.Decrypt(ConfigurationManager.AppSettings["ConnectionString"].ToString(), true)))
                            {
                                conn.Open();
                                conn.ChangeDatabase(moduleDatabase);
                                using (SqlCommand cmd = new SqlCommand("SP_PROCS_INSIDER_CONFIG_SMTP", conn))
                                {
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.CommandTimeout = 0;
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.Add(new SqlParameter("@Mode", "GET_Smtp_Config_List"));
                                    cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                                    cmd.Parameters.Add(new SqlParameter("@COMPANY_ID", companyId));
                                    SqlDataReader rdr = cmd.ExecuteReader();
                                    if (rdr.HasRows)
                                    {
                                        while (rdr.Read())
                                        {
                                            isSMTPConfigExists = true;
                                            defaultEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["DEFAULT_EMAIL"]))) ? Convert.ToString(rdr["DEFAULT_EMAIL"]) : String.Empty;
                                            disclosureEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]))) ? Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]) : String.Empty;
                                            smtpHostName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_HOST_NAME"]))) ? Convert.ToString(rdr["SMTP_HOST_NAME"]) : String.Empty;
                                            port = Convert.ToInt32(rdr["PORT"]);
                                            ssl = (!String.IsNullOrEmpty(Convert.ToString(rdr["SSL"]))) ? (Convert.ToString(rdr["SSL"]) == "Yes" ? true : false) : false;
                                            smtpUserName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_USER_NAME"]))) ? Convert.ToString(rdr["SMTP_USER_NAME"]) : String.Empty;
                                            password = (!String.IsNullOrEmpty(Convert.ToString(rdr["PASSWORD"]))) ? CryptorEngine.Decrypt(Convert.ToString(rdr["PASSWORD"]), true) : String.Empty;
                                            userDefaultCredential = (!String.IsNullOrEmpty(Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]))) ? (Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]) == "Yes" ? true : false) : false;
                                        }
                                    }
                                    rdr.Close();
                                }
                                conn.Close();
                            }

                            if (isSMTPConfigExists)
                            {
                                APICommittee objCommittee = committeeResponse.APICommitteeList[0];
                                APIMeeting objMeeting = objCommittee.listMeeting[0];
                                APIAgendaItems objAgenda = objMeeting.agendaItems[0];
                                List<APIUser> lisUser = objMeeting.listMeetingMembers;

                                String layoutTemplate = String.Empty;
                                String subject = "Comments Shared Email";
                                foreach (APIUser objUser in lisUser)
                                {
                                    layoutTemplate = String.Empty;
                                    layoutTemplate += "<p>Hi <b>" + objUser.UserName + "</b>,</p>";
                                    layoutTemplate += "<br>";
                                    layoutTemplate += "<p>";
                                    layoutTemplate += "Comments has been shared for Committee <b>" + objCommittee.committeeName + "</b> for Meeting <b>" + objMeeting.meetingTitle + "</b> for Item <b>" + objAgenda.agendaTitle + "</b> by " + objCommittee.createdBy + ".";
                                    layoutTemplate += "</p>";
                                    layoutTemplate += "<br>";
                                    layoutTemplate += "<p> Thanks &amp; Regards,</p>";
                                    layoutTemplate += "<br>";
                                    layoutTemplate += "<p>" + objCommittee.createdBy + "</p>";

                                    EmailHelper.SendMail(defaultEmail, objUser.Email, subject, layoutTemplate, null, smtpHostName, ssl, smtpUserName, password, userDefaultCredential, port, objCommittee.createdBy);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "EmailSender", "SendCommentsShareEmail", "Admin", 1, companyId);
            }
        }

        #endregion

        #region "Create Draft MOM Task and Send Email to each User"

        private static void CreateDraftMOMTaskAndSendEmail(Int32 companyId, Int32 committeeId, Int32 meetingId, Int32 draftMOMId, String userLogin, String dataBase, Int32 userId, String mode)
        {
            List<String> listAttachments = null;

            try
            {
                //String defaultEmail = String.Empty, disclosureEmail = String.Empty, smtpHostName = String.Empty, smtpUserName = String.Empty, password = String.Empty;
                //Int32 port = 0;
                //bool ssl = false, userDefaultCredential = false, isSMTPConfigExists = false;
                EmailRepository emailRepo = new EmailRepository();
                EmailResponse emailRes = new EmailResponse();
                Email objEmail = null;

                //using (SqlConnection conn = new SqlConnection(CryptorEngine.Decrypt(ConfigurationManager.AppSettings["ConnectionString"].ToString(), true)))
                //{
                //    conn.Open();
                //    conn.ChangeDatabase(dataBase);
                //    using (SqlCommand cmd = new SqlCommand("SP_PROCS_INSIDER_CONFIG_SMTP", conn))
                //    {
                //        cmd.CommandType = CommandType.StoredProcedure;
                //        cmd.CommandTimeout = 0;
                //        cmd.Parameters.Clear();
                //        cmd.Parameters.Add(new SqlParameter("@Mode", "GET_Smtp_Config_List"));
                //        cmd.Parameters.Add(new SqlParameter("@SET_COUNT", SqlDbType.Int)).Direction = ParameterDirection.Output;
                //        cmd.Parameters.Add(new SqlParameter("@COMPANY_ID", companyId));
                //        SqlDataReader rdr = cmd.ExecuteReader();
                //        if (rdr.HasRows)
                //        {
                //            while (rdr.Read())
                //            {
                //                isSMTPConfigExists = true;
                //                defaultEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["DEFAULT_EMAIL"]))) ? Convert.ToString(rdr["DEFAULT_EMAIL"]) : String.Empty;
                //                disclosureEmail = (!String.IsNullOrEmpty(Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]))) ? Convert.ToString(rdr["CONTINUAL_DISCLOSURE_EMAIL"]) : String.Empty;
                //                smtpHostName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_HOST_NAME"]))) ? Convert.ToString(rdr["SMTP_HOST_NAME"]) : String.Empty;
                //                port = Convert.ToInt32(rdr["PORT"]);
                //                ssl = (!String.IsNullOrEmpty(Convert.ToString(rdr["SSL"]))) ? (Convert.ToString(rdr["SSL"]) == "Yes" ? true : false) : false;
                //                smtpUserName = (!String.IsNullOrEmpty(Convert.ToString(rdr["SMTP_USER_NAME"]))) ? Convert.ToString(rdr["SMTP_USER_NAME"]) : String.Empty;
                //                password = (!String.IsNullOrEmpty(Convert.ToString(rdr["PASSWORD"]))) ? CryptorEngine.Decrypt(Convert.ToString(rdr["PASSWORD"]), true) : String.Empty;
                //                userDefaultCredential = (!String.IsNullOrEmpty(Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]))) ? (Convert.ToString(rdr["USER_DEFAULT_CREDENTIAL"]) == "Yes" ? true : false) : false;
                //            }
                //        }
                //        rdr.Close();
                //    }

                //if (isSMTPConfigExists)
                //{
                
                MeetingRepository meetingRepo = new MeetingRepository();               
                DraftMOMTask objTask = null;
                if (mode.ToUpper() == "REPLACEDRAFT")
                {
                    objTask = new DraftMOMTask();
                    objTask.companyId = companyId;
                    objTask.committee = new Committee
                    {
                        ID = committeeId
                    };
                    objTask.meeting = new Meeting
                    {
                        ID = meetingId
                    };
                    objTask.moduleDatabase = dataBase;
                    objTask.mode = mode;
                    objTask.draftMOMId = draftMOMId;
                    meetingRepo.DeleteDraftMOMTask(objTask);
                }

                List<User> lstUsers = meetingRepo.GetDraftMOMUsers(draftMOMId, userLogin, companyId, dataBase, userId);
                new LogHelper().AddExceptionLogs("before user", "", "", "EmailSender", "CreateDraftMOMTaskAndSendEmail", userLogin, 1, companyId);
                if (lstUsers != null)
                {
                    if (lstUsers.Count > 0)
                    {
                        new LogHelper().AddExceptionLogs("AFTER user", "", "", "EmailSender", "CreateDraftMOMTaskAndSendEmail", userLogin, 1, companyId);
                        foreach (User objUser in lstUsers)
                        {
                            objTask = new DraftMOMTask();
                            objTask.companyId = companyId;
                            objTask.committee = new Committee
                            {
                                ID = committeeId
                            };
                            objTask.meeting = new Meeting
                            {
                                ID = meetingId
                            };
                            objTask.draftMOMId = draftMOMId;
                            objTask.assignedToUser = new User
                            {
                                ID = objUser.ID,
                                emailId = objUser.emailId
                            };
                            objTask.createdBy = userLogin;
                            objTask.moduleDatabase = dataBase;
                            objTask.mode = mode;
                            new LogHelper().AddExceptionLogs("BEFORE", "", "", "EmailSender", "CreateDraftMOMTaskAndSendEmail", userLogin, 1, companyId);
                            meetingRepo.CreateDraftMOMTask(objTask);
                            new LogHelper().AddExceptionLogs("AFTER", "", "", "EmailSender", "CreateDraftMOMTaskAndSendEmail", userLogin, 1, companyId);
                            if (objTask.isSendEmail)
                            {
                                listAttachments = new List<String>();
                                listAttachments = meetingRepo.GetDraftMOMAttachments(objTask);

                                objEmail = new Email();
                                objEmail.moduleDatabase = dataBase;
                                objEmail.taskName = "Draft MOM Approval";
                                objEmail.isMeeting = true;
                                objEmail.companyId = companyId;
                                objEmail.meeting = new Meeting
                                {
                                    ID = meetingId
                                };
                                objEmail.user = new User
                                {
                                    ID = objUser.ID
                                };
                                emailResponse = emailRepo.GetTemplateHeaderDetailByNameAndCompanyId(objEmail);
                                if (emailResponse.StatusFl)
                                {
                                    if (emailResponse.EmailList != null)
                                    {
                                        if (emailResponse.EmailList.Count > 0)
                                        {
                                            objEmail = new Email();
                                            objEmail = emailResponse.EmailList[0];
                                            EmailHelper.SendDraftMOMTaskEmail(objEmail.subject, objEmail.layoutTemplate, listAttachments, userLogin, objTask);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                //}
                //conn.Close();
                //}
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.Trim(), ex.Source, ex.StackTrace, "EmailSender", "CreateDraftMOMTaskAndSendEmail", userLogin, 1, companyId);
            }
        }

        #endregion

        /*
        public static string ConvertToVCalendarDateString(DateTime d)
        {
            d = d.ToUniversalTime();
            string yy = d.Year.ToString();
            string mm = d.Month.ToString("D2");
            string dd = d.Day.ToString("D2");
            string hh = d.Hour.ToString("D2");
            string mm2 = d.Minute.ToString("D2");
            string ss = d.Second.ToString("D2");
            string s = yy + mm + dd + "T" + hh + mm2 + ss + "Z"; // Pass date as vCalendar format YYYYMMDDTHHMMSSZ (includes middle T and final Z) '
            return s;
            
            MailMessage mm = new MailMessage();
            mm.From = new MailAddress("procstest@gmail.com");
            mm.Subject = "Meeting Scheduled";
            mm.To.Add(new MailAddress("amitaverma04@gmail.com", "Romil Nehru"));

            //mm.Body = "Please Attend the meeting with this schedule";
            mm.Body = "<html><body><p>Respected Sir/Madam2,</p><p>This is to inform you that 1st Meeting of Audit Committee of Audit Committee has been scheduled.</p><p>The venue of the meeting is at Delhi,Board Room, 3rd floor, ACME Corp Ltd, DLF Promnegate, Delhi on Jul 17, 2020 from 08:15 AM to 10:15 AM.</p><p>Thanks &amp; Regards,</p><p>Abhinav</p></body></html>";

            StringBuilder str = new StringBuilder();
            str.AppendLine("BEGIN:VCALENDAR");
            str.AppendLine("PRODID:-//Schedule a Meeting");
            str.AppendLine("VERSION:2.0");
            str.AppendLine("METHOD:REQUEST");
            str.AppendLine("BEGIN:VEVENT");

            String stDate = DateTime.Now.AddDays(2).ToString("yyyy-MM-dd");

            String STTIME = "8:15:00 AM";
            String EDTIME = "10:15:00 AM";

            DateTime st = Convert.ToDateTime(stDate + " " + STTIME);
            DateTime et = Convert.ToDateTime(stDate + " " + EDTIME);

            //Set the start time of meeting
            //str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddHours(2.0)));
            str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", ConvertToVCalendarDateString(st)));

            //Set the end time of meeting
            //str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddHours(5.0)));
            str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", ConvertToVCalendarDateString(et)));

            //Set the location of meeting
            str.AppendLine("LOCATION: Pioneer House, Sector 16, Noida");
            str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            str.AppendLine(string.Format("DESCRIPTION:{0}", mm.Body));
            str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", mm.Body));
            str.AppendLine(string.Format("SUMMARY:{0}", mm.Subject));
            str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", mm.From.Address));

            //Loop to show the attendees
            foreach (MailAddress ma in mm.To)
            {
                str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", ma.DisplayName, ma.Address));
            }

            str.AppendLine("STATUS:CONFIRMED");
            str.AppendLine("BEGIN:VALARM");
            str.AppendLine("TRIGGER:-PT15M");
            str.AppendLine("ACTION:Accept");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("END:VALARM");
            str.AppendLine("END:VEVENT");
            str.AppendLine("END:VCALENDAR");

            mm.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "smtp.gmail.com";
            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new System.Net.NetworkCredential("procstest@gmail.com", "P@ssw0rd@123");
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;

            var htmlContentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Text.Html);
            var avHtmlBody = AlternateView.CreateAlternateViewFromString(mm.Body, htmlContentType);
            mm.AlternateViews.Add(avHtmlBody);

            System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
            contype.Parameters.Add("method", "REQUEST");
            contype.Parameters.Add("name", "Meeting.ics");
            AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
            mm.AlternateViews.Add(avCal);

            smtp.Port = 587;
            smtp.Send(mm);
            return;
           
    }
     */

    }
}
