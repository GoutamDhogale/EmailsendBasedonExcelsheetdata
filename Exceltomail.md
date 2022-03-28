using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Syncfusion.XlsIO;
using GemBox.Spreadsheet;
using Nancy.Json;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Globalization;

namespace ReadXlsxFile
{
    class Program
    {
        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");          //Excel Reading and apply consition

            var workbook = ExcelFile.Load("Name.xlsx");   //Text1

            // Select the first worksheet from the file.
            var worksheet = workbook.Worksheets[0];

            // Create DataTable from an Excel worksheet.
            var dataTable = worksheet.CreateDataTable(new CreateDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0,
                NumberOfColumns = 20,
                NumberOfRows = worksheet.Rows.Count ,
                Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
            });

            //// Write DataTable content
            var sb = new StringBuilder();
            foreach (DataRow row in dataTable.Rows)
            {
                sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}", row[0], row[2], row[6], row[8], row[19]);
                var name = row[8];
                var tkt1 = row[0];
                var priority = row[2];
                var Email = "YourMailId";
                var ticketNumber = Convert.ToString(tkt1);
                var AssignedTo = Convert.ToString(name);
                

                var a = row[6];
                
                //var jsonStringName = new JavaScriptSerializer();
                //var jsonStringResult = jsonStringName.Serialize(a);
                var TicketOpened = Convert.ToDateTime(a);
                //string MyString = jsonStringResult.ToString();
                //DateTime dt1 = DateTime.ParseExact(MyString, "dd-MM-yyyy ",
                //                                  null);

                //long mylong = (long)a;
                //var today = DateTime.Now.ToString("d/M/yyyy");
                //var b = long.Parse(today);
                var today = DateTime.Now;
                //long hel = (long)today;
                //get difference of two dates
                var NumberOfDays = (today - TicketOpened).Days;                         //Number of day calculation
                if (NumberOfDays > 29)
                {
                    try
                    {                                                             //Email Part
                        // Create the Outlook application.
                        Outlook.Application oApp = new Outlook.Application();
                        // Create a new mail item.
                        Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                        // Set HTMLBody. 
                        //add the body of the email
                         
                        oMsg.HTMLBody = "Hello " + AssignedTo + ",<br/>You have been Tag to this ticket " + NumberOfDays+ " days ago.Please complete this ticket as fast as possible.<br/> Ticket Number:" + ticketNumber+ " <br/> Ticket Opend at: " + TicketOpened+"<br/>Priority:"+ priority + "<br/><br/>Regards,<br/>abc";
                      
                        //Subject line
                        oMsg.Subject = "Ticket Status";
                        // Add a recipient.
                        Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                        // Change the recipient in the next line if necessary.
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(Email);
                      
                        oRecip.Resolve();
                        // Send.
                        oMsg.Send();
                        // Clean up.
                        oRecip = null;
                        oRecips = null;
                        oMsg = null;
                        oApp = null;
                    }//end of try block
                    catch (Exception ex)
                    {
                        throw ex;
                    }//end of catch
                }
                Console.WriteLine(NumberOfDays);
            }
           

        }

       
    }
}
    
