using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace splitPdfApp
{
    class Program
    {
        //NominalDataEntities DBAccess = new NominalDataEntities();
         static string globalIppis = "";
        
        static void Main(string[] args)
        {
            //splitPDFUp();
            //dbaseAccess();
            MyMain();
        }

        public static void MyMain()
        {                        
            for (int i = 9005; i <= 9288; i++)
            {             

                    String ippis_result = "E:/Projects/Split-PDF/result/" + i + ".pdf";
                    string newresult = ParsePdf(ippis_result);
                    //string staffEmail = dbaseAccess(newresult);
                    //globalIppis = newresult;

                if (newresult != "" || newresult != "Oracl")
                {
                    string staffEmail = dbaseAccess(newresult, "email");
                    string staffName = dbaseAccess(newresult, "name");
                    globalIppis = newresult;
                    if (!string.IsNullOrEmpty(staffEmail))
                    {
                        bool testEmail = IsValidEmail(staffEmail);
                        if (testEmail)
                        {

                            try
                            {
                                File.Move("E:/Projects/Split-PDF/result/" + i + ".pdf", "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");
                            }
                            catch
                            {

                            }


                            Console.WriteLine(newresult);
                            string mailMessage = staffName + "<br />" + " Find attached document for June 2019 payslip ....Finance and Account - UCH, Ibadan." + "<br />"+ "<br />"+ "Powered By Information Technology Department";
                            string mailSubject = "UCH: Payslip June 2019 for IPPIS Number "+ newresult;
                            //Console.WriteLine(staffEmail);
                            //Console.ReadLine();
                            //Send("payslip@uch-ibadan.org.ng", "payslipPWD123!", staffEmail, "Payslip" + " " + newresult, "Payslip" + " " + newresult, "mail.uch-ibadan.org.ng", 22, "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");
                            //Send("payslip@uch-ibadan.org.ng", "payslipPWD123!", "idowu.ogedengbe@gmail.com", mailMessage, "Payslip", "mail.uch-ibadan.org.ng", 22, "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");
                            Send("payslip@uch-ibadan.org.ng", "payslipPWD123!", staffEmail, mailMessage, mailSubject, "mail.uch-ibadan.org.ng", 22, "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");
                            //Console.ReadLine();
                        }
                        else
                        {
                            updateDB(globalIppis, "Invalid Email");
                            Console.WriteLine("Invalid Email");
                            //Console.ReadLine();
                        }
                    }
                    else
                    {
                        updateDB(globalIppis, "No Email");
                        Console.WriteLine("No Email");
                        //Console.ReadLine();
                    }
                }


                /*try
                {
                    File.Move("E:/Projects/Split-PDF/result/" + i + ".pdf", "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");
                }
                catch
                {

                }


                Console.WriteLine(newresult);
                Console.ReadLine();
                Send("joseph.jolaosho@uch-ibadan.org.ng", "jossyPWD123!", "jajossy@yahoo.com", "Payslip"+" "+newresult, "Payslip"+" "+newresult, "mail.uch-ibadan.org.ng", 22, "E:/Projects/Split-PDF/ippisresult/" + newresult + ".pdf");*/

                //} 
                Console.WriteLine(i);
            }
            Console.ReadLine();
        }

        public static string ParsePdf(string filename)
        //public static int ParsePdf(string filename)
        {            
            if (!File.Exists(filename))
                throw new FileNotFoundException("fileName");

            using (PdfReader textreader = new PdfReader(filename))
            {
                StringBuilder sb = new StringBuilder();

                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                for (int page = 0; page < textreader.NumberOfPages; page++)
                {
                    string text = PdfTextExtractor.GetTextFromPage(textreader, page + 1, strategy);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        sb.Append(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(text))));
                    }
                }
                string sb_final = sb.ToString();               
                int new_file_name_index = sb_final.IndexOf("Number");
                int startValue = new_file_name_index + 8;
                string new_file_name = sb_final.Substring(startValue, 6);

                return new_file_name;

                // Code adjustment

                /*if (new_file_name == "201363")

                {
                    return new_file_name;
                }
                else
                {
                    return "";
                }*/               
            }
        }        

        public static void Send(string from, string password, string to, string Message, string subject, string host, int port, string file)
        {

            MailMessage email = new MailMessage();
            email.From = new MailAddress(from);
            email.To.Add(to);
            email.Subject = subject;
            email.Body = Message;
            SmtpClient smtp = new SmtpClient(host, port);
            smtp.UseDefaultCredentials = false;
            NetworkCredential nc = new NetworkCredential(from, password);
            smtp.Credentials = nc;
            smtp.EnableSsl = true;
            
            email.IsBodyHtml = true;
            email.Priority = MailPriority.Normal;
            email.BodyEncoding = Encoding.UTF8;

            if (file.Length > 0)
            {
                Attachment attachment;
                attachment = new Attachment(file);
                email.Attachments.Add(attachment);
            }

            // smtp.Send(email);
            smtp.SendCompleted += new SendCompletedEventHandler(SendCompletedCallBack);
            string userstate = "sending ...";
            smtp.SendAsync(email, userstate);
            
        }

        private static void SendCompletedCallBack(object sender, AsyncCompletedEventArgs e)
        {
            string result = "";
            if (e.Cancelled)
            {
                //MessageBox.Show(string.Format("{0} send canceled.", e.UserState), "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Console.WriteLine(string.Format("{0} send canceled.", e.UserState), "Message");
                updateDB(globalIppis, "Cancelled");
            }
            else if (e.Error != null)
            {
                //MessageBox.Show(string.Format("{0} {1}", e.UserState, e.Error), "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Console.WriteLine(string.Format("{0} {1}", e.UserState, e.Error), "Message");
                updateDB(globalIppis, "Not Sent");
            }
            else
            {
                //MessageBox.Show("your message is sended", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Console.WriteLine(string.Format("{0} {1}", "your message is sended", "Message"));
                updateDB(globalIppis, "Sent");
            }

        }

        private static string dbaseAccess(string ippis, string token)
        //public static void dbaseAccess()
        {
            //IEnumerable<StaffData> query = DBAccess.StaffDatas.Select();
            string finalValue;
            using (var context = new NominalDataEntities())
            {
                var stdQuery = (from d in context.StaffDatas
                                select new
                                {
                                    Id = d.Id,
                                    PinNo = d.PinNo,
                                    IppisNo = d.IppisNo,
                                    Fullname = d.Fullname,
                                    EmailAddress = d.EmailAddress,
                                    Department = d.Department,
                                    Comment = d.Comment
                                }).Where(x => x.IppisNo == ippis).FirstOrDefault();

                /*foreach (var q in stdQuery)
                {
                    Console.WriteLine("PinNo : " + q.PinNo + ", IppisNo : " + q.IppisNo + ", Department : " + q.Department);
                }*/

                
                //Console.ReadLine();

                if (stdQuery != null && token == "email")
                    finalValue = stdQuery.EmailAddress;
                else if(stdQuery != null && token == "name")
                    finalValue = stdQuery.Fullname;
                else
                finalValue = null;

            }
            
            return finalValue;
        }

        public static void updateDB(string ippis, string message)
        {
            using (var context2 = new NominalDataEntities())
            {
                //var result = context2.StaffDatas.SingleOrDefault(b => b.IppisNo == ippis);
                var result = context2.StaffDatas.FirstOrDefault(b => b.IppisNo == ippis);
                if (result != null)
                {
                    result.Comment = message;
                    context2.SaveChanges();
                }
            }
        }

        public static bool IsValidEmail(string strIn)
        {
            // Return true if strIn is in valid e-mail format.
            return Regex.IsMatch(strIn, @"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");
        }

        public static void splitPDFUp()
        {
            //variables
            String source_file = "E:/Projects/Split-PDF/source/JUNE_2019.pdf";
            String result = "E:/Projects/Split-PDF/result/";

            PdfCopy copy;

            //create PdfReader object
            PdfReader reader = new PdfReader(source_file);

            for (int i = 1; i <= reader.NumberOfPages; i++)           
            {
                //if (i == 3546)
                //{
                //create Document object
                Document document = new Document();
                copy = new PdfCopy(document, new FileStream(result + i + ".pdf", FileMode.Create));
                //open the document 
                document.Open();
                //add page to PdfCopy 
                copy.AddPage(copy.GetImportedPage(reader, i));
                //close the document object
                document.Close();
                Console.WriteLine(i);

            }
        }
    }
}
