using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Reflection;

namespace AutoMail
{
    class Program
    {
        string smtpAddress = "";
        int portNumber = 0;
        bool enableSSL = true;
        string emailFrom = "";
        string password = "";
        string emailTo = "";
        string subject = "";
        string body = "";
        string year = DateTime.Now.Year.ToString();
        string monthnumber = DateTime.Now.ToString("yyyy-MM-MMM");
        string currentdate = DateTime.Now.ToString("yyyy-MM-dd");
        string fileName = "";
        string filePath = "";
        int fileCount = 0;
        int runCount = 0;
        static void Main(string[] args)
        {
            Program p = new Program();
            p.ReadConfig();
            p.sendMail();
            

        }

        //Read config file
    
        public void ReadConfig()
        {
            var data = File
             .ReadAllLines(@"D:\Amit Programs\Automatic mail sending\AutoMail\Parameter\parameter.txt")
             .Select(x => x.Split('='))
             .Where(x => x.Length > 1)
             .ToDictionary(x => x[0].Trim(), x => x[1]);
            smtpAddress = data["smtpAddress"].ToString().Trim() ;
            portNumber =Convert.ToInt32(data["portName"].ToString().Trim());
            emailFrom = data["emailFrom"].ToString().Trim();
            password = data["password"].ToString().Trim();
            emailTo = data["emailTo"].ToString().Trim();
            subject = data["subject"].ToString().Trim();
            body = data["body"].ToString().Trim();
            filePath = data["pdfPath"].ToString().Trim()+"\\"+year+"\\"+monthnumber+"\\"+currentdate+"\\"+"report";
         // filePath = data["pdfPath"].ToString().Trim();
         
        }
        //method to send mail
        public void sendMail()
        { 
            
           
            LogWriter logwrite = new LogWriter();
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(emailFrom);
                mail.To.Add(emailTo);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true;
                // Can set to false, if you are sending pure text.

                if (Directory.Exists(filePath))
                {
                    if (System.IO.Directory.GetFiles(filePath, "*", SearchOption.AllDirectories).Length != 0)
                    {
                        try
                        {
                            foreach (FileInfo file in new DirectoryInfo(filePath ).GetFiles("*.pdf"))
                            {
                                mail.Attachments.Add(new Attachment(file.FullName));
                                fileName = fileName + file.FullName.ToString() + Environment.NewLine;
                                fileCount = fileCount + 1;
                            }
                            
                          
                                                    
                            using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                            {
                                runCount = runCount + 1;
                                smtp.Credentials = new NetworkCredential(emailFrom, password);
                                smtp.EnableSsl = enableSSL;
                                smtp.Send(mail);
                            }

                            
                            logwrite.LogWrite("Total program run count="+runCount+Environment.NewLine + fileCount + " files sucessfully sent"+Environment.NewLine+"Below mentioned files are sent sucessfully sent" + Environment.NewLine + fileName + Environment.NewLine + "Mail sent complete !!! ");
                           
                            fileCount = 0;
                            fileName = "";
                            Environment.Exit(-1);
                        }
                        catch (Exception e)
                        {
                            logwrite.LogWrite(e.ToString());
                            // secound attempt to run the program
                            if (runCount <= 2)
                            {
                                sendMail();
                            }
                            else
                            {
                                logwrite.LogWrite("Total run count="+ runCount);
                                runCount = 0;
                                Environment.Exit(-1);
                            }
                                                        
                        }
                    }
                    else
                    {
                       
                        logwrite.LogWrite("Folder donot contain any files");
                        Environment.Exit(-1);
                    }
                }

                else
                {
                    logwrite.LogWrite("Folder not found");
                    Environment.Exit(-1);
                }
            }


        }

    }
    public class LogWriter
    {
        private string m_exePath = string.Empty;
        //public LogWriter(string logMessage)
        //{
        //    LogWrite(logMessage);
        //}
        public void LogWrite(string logMessage)
        {
            string path = @"D:\mailtest";
            m_exePath = path;
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
               // txtWriter.WriteLine("  :");
                txtWriter.WriteLine(logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
            }
        }
    }
    }


