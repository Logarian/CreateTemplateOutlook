using System.Text;
using MsOutlook = Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
using System.IO;
//using System.Net;
//using System.Net.Mail;

namespace TestCreateTemplateOutlook
{
    class Program
    {
        static void Main(string[] args)
        {

            using (var reader = new StreamReader(@"Список.csv", Encoding.Default))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] elementsLine = line.Split('#');

                    OutlookApp outlookApp = new OutlookApp();
                    MsOutlook.MailItem mailItem = outlookApp.CreateItem(MsOutlook.OlItemType.olMailItem);
                    mailItem.Subject = "Площадка №" + elementsLine[0].Replace("38.","");

                    mailItem.To = elementsLine[3];

                    string[] forBlanks = elementsLine[0].Split('.');
                    if (forBlanks[1].Length == 1)
                    {
                        MsOutlook.Attachment attachment = mailItem.Attachments.Add(@"C:\Путь к бланкам\" + forBlanks[0] + "_00000" + forBlanks[1] + ".pdf");
                    } else if (forBlanks[1].Length == 2)
                    {
                        MsOutlook.Attachment attachment = mailItem.Attachments.Add(@"C:\Путь к бланкам\" + forBlanks[0] + "_0000" + forBlanks[1] + ".pdf");
                    } else if (forBlanks[1].Length == 3)
                    {
                        MsOutlook.Attachment attachment = mailItem.Attachments.Add(@"C:\Путь к бланкам\" + forBlanks[0] + "_000" + forBlanks[1] + ".pdf");
                    }

                    MsOutlook.Attachment attachment2 = mailItem.Attachments.Add(@"Какое-либо вложение.rar");

                    mailItem.HTMLBody = "<html><body>" +
                        "<p>Тело письма</p>" +
                        "</body></html>";
                    //mailItem.Display(false);
                    mailItem.SaveAs("C:\\Users\\<...>\\Documents\\Площадка_" + elementsLine[0].Replace(".","_") + ".oft", MsOutlook.OlSaveAsType.olTemplate);
                    
                }
            }
        }
    }
}
