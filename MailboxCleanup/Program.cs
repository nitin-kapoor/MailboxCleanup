using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Exchange.WebServices.Data;


namespace ExchangeMailbox
{
    class Program
    {
        public static string mailuser = string.Empty;
        public static string mailpwd = string.Empty;
        public static string maildomain = string.Empty;
        public static string mailservice = string.Empty;
        public static string mailserver = string.Empty;
        public static string strApproveMail;
        public static string strRejectMail;
        public static string LogFile;
        string strAppFolderID;
        string strRejFolderID;
        string strIgnFolderID;

        //This function will load all the Exchange configuration/credentials from Settings.xml
        public static void LoadSettings()
        {
            string fileName = "Settings.xml";
            
            if(File.Exists(fileName)) 
            {
                XDocument xmlDoc = XDocument.Load(fileName);

                mailuser = xmlDoc.Root.Element("mail_user").Value;
                mailpwd = xmlDoc.Root.Element("mail_pwd").Value;
                maildomain = xmlDoc.Root.Element("mail_domain").Value;
                mailservice = xmlDoc.Root.Element("mail_service").Value;
                mailserver = xmlDoc.Root.Element("mail_server").Value;
                LogFile = xmlDoc.Root.Element("LogFile").Value;
                strApproveMail = xmlDoc.Root.Element("SubjectLineApproved").Value;
                strRejectMail = xmlDoc.Root.Element("SubjectLineRejected").Value;

                WriteLog("Loading Process Settings");
                WriteLog("Process Settings loaded successfully");
            }
            else
            {
                WriteLog("Process Settings not found. Aborting process");
                Environment.Exit(0);
            }
        }

        public static void WriteLog(string log) 
        {
            if (log == "")
            {
                Console.WriteLine("\n");
                File.AppendAllText(LogFile, Environment.NewLine);
            }
            else
            {
                Console.WriteLine(DateTime.Now.ToString() + ":      " + log);
                File.AppendAllText(LogFile, DateTime.Now.ToString() + ":    " + log + Environment.NewLine);
            }
        }

        public void Start()
        {
            LoadSettings();
            WriteLog("Job Started");
            WriteLog("Creating EWS Connection");
            ExchangeService service = new ExchangeService();
            service.Timeout = 600000;
            service.Credentials = new NetworkCredential(mailuser, mailpwd, maildomain);
            service.Url = new Uri(mailservice);
            WriteLog("EWS Connection Created");
            
            WriteLog("Creating Event Subscription");
            WriteLog("Event Subscription Started");
            
            ItemView view = new ItemView(100);
            FindItemsResults<Item> findResults;

            WriteLog("Checking Existing Emails");
            do
            {
                //To check count of emails in Junk folder.
                findResults = service.FindItems(WellKnownFolderName.JunkEmail, view);

                FolderView view2 = new FolderView(100);
                view2.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                view2.PropertySet.Add(FolderSchema.DisplayName);
                view2.Traversal = FolderTraversal.Deep;

                FindFoldersResults folderResult = service.FindFolders(WellKnownFolderName.Inbox, view2);
                //To find specific folder
                foreach(Folder f in folderResult) 
                {
                    switch(f.DisplayName)
                    {
                        case "Approved":
                            strAppFolderID = f.Id.ToString();
                            break;
                        case "Rejected":
                            strRejFolderID = f.Id.ToString();
                            break;
                        case "Ignored":
                            strIgnFolderID = f.Id.ToString();
                            break;
                    }
                }

                int Total = findResults.Count();
                WriteLog("There are " + Total + " Emails in Junk Folder");
                WriteLog("");
                if (Total > 0)
                {
                    WriteLog("========== Mailbox Cleanup Started ==========");
                    WriteLog("");
                    Item last = findResults.Last();
                    foreach (Item item in findResults.Items)
                    {
                        EmailMessage email = EmailMessage.Bind(service, item.Id);
                        if (checkNullString(email.Subject) != "")
                        {
                            if (email.Subject.Contains(strApproveMail) || email.Subject.Contains(strRejectMail))
                            {
                                string strSubject = email.Subject.Trim();
                                WriteLog(strSubject);
                                string[] strSubjectSplit = strSubject.Split('|');

                                WriteLog("Split One Subject : " + strSubjectSplit[0]);
                                WriteLog("Split Two Subject : " + strSubjectSplit[1]);
                                WriteLog("Split Three Subject : " + strSubjectSplit[2]);

                                switch (strSubjectSplit[1])
                                {
                                    case "APPROVED":
                                        FolderId ApprovedFolderID = new FolderId(strAppFolderID);
                                        item.Move(ApprovedFolderID);
                                        WriteLog("Email Moved to [APPROVED] Folder");
                                        WriteLog("");
                                        break;

                                    case "REJECTED":
                                        FolderId RejectFolderId = new FolderId(strRejFolderID);
                                        item.Move(RejectFolderId);
                                        WriteLog("Email Moved to [REJECTED] Folder");
                                        WriteLog("");
                                        break;
                                }
                            }
                            else
                            {
                                string strSubject = email.Subject.Trim();
                                FolderId IgnoredFolderId = new FolderId(strIgnFolderID);
                                item.Move(IgnoredFolderId);
                                WriteLog(strSubject);
                                WriteLog("Subject Line doesn't contain APPROVED or REJECTED keyword. Email moved to [IGNORED] Folder");
                                WriteLog("");
                            }
                        }
                        else
                        {
                            FolderId IgnoredFolderId = new FolderId(strIgnFolderID);
                            item.Move(IgnoredFolderId);
                            WriteLog("Subject Line is Empty. Email moved to [IGNORED] Folder");
                            WriteLog("");
                        }
                        
                        if(item.Equals(last))
                        {
                            WriteLog("========== Mailbox Cleanup Ended ==========");
                            WriteLog("");
                        }
                    }
                }
                else
                {
                    WriteLog("Mailbox Cleanup is not required because there are no emails in Junk folder.");
                }

            } while (findResults.MoreAvailable);

        }

        public string checkNullString(string mystring)
        {
            string myReturn = mystring ?? "";
            return myReturn.Trim();
        }

        static void Main(string[] args)
        {
            Program phase1 = new Program();
            phase1.Start();
            Console.ReadLine();
        }
    }
}
