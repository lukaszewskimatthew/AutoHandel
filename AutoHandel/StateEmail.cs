using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace AutoHandel
{
    class StateEmail
    {
        string sendEmail;
        string distract;
        string gradeLevel;
        string body;

        public StateEmail(string inEmail, string inDist, string inGrade)
        {
            sendEmail = inEmail;
            distract = inDist;
            gradeLevel = inGrade;
        }

        public void AddToBody(string addStr) { body += addStr; }

        public void SendEmail(string uName, string passWord)
        {
            MailMessage mMess = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            try
            {
                mMess.From = new MailAddress("iscreports@wflboces.org", "Suspension Center");
                mMess.To.Add(new MailAddress(sendEmail));
                mMess.Subject = DateTime.Now.ToString("MM/dd/yyyy") + " Attendance";
                mMess.Body = body;
                mMess.BodyEncoding = Encoding.UTF8;
                mMess.SubjectEncoding = Encoding.UTF8;

                smtp.Host = "smtp-relay.gmail.com";
                smtp.Port = 587;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("iscreports@wflboces.org", "G1nger#7");
                //smtp.EnableSsl = true; (Not sure why this makes the connection drop)
                smtp.Send(mMess);
            }

            catch (SmtpException ex) { throw new SmtpException(ex.Message); }

            finally
            {
                mMess.Dispose();
                smtp.Dispose();
            }
        }

        public string ContactEmail { get { return sendEmail; } }

        public string Distract { get { return distract; } }

        public string GradeLevel { get { return gradeLevel; } }

        public string Body { get { return body; } }
    }
}
