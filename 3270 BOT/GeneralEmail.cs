using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace _3270_BOT
{
    class GeneralEmail
    {
        string table;
        string body;
        string add;

        public void sendEmail(string subject, string cc, string to, string body)
        {


            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            //mailItem.SentOnBehalfOfName = "CCCmisqa";

            mailItem.To = to;
            //mailItem.To = "ccc.rhbway@nonSMTP.rhbgroup.com";
            mailItem.CC = cc;




            //mailItem.DeleteAfterSubmit = true;
            mailItem.HTMLBody = body;
            //mailItem.Attachments.Add(logPath);//logPath is a string holding path to the log.txt file
            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Send();

            table = "";
            body = "";

        }





    }
}
