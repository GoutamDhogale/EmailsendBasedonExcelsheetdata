openFileDialog1.ShowDialog();
            string fileName = openFileDialog1.FileName;
            string dbf_Path = System.IO.Path.GetFileName(fileName);

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");          //Excel Reading and apply consition
            var workbook = ExcelFile.Load(fileName);                         //"Test1.xlsx"
            //DateTime creation = File.GetCreationTime("Test1.xlsx");
            //Console.WriteLine(creation);
            //if (creation == DateTime.Today)
            //{
            // Select the first worksheet from the file.
            var worksheet = workbook.Worksheets[0];

            // Create DataTable from an Excel worksheet.
            var dataTable = worksheet.CreateDataTable(new CreateDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0,
                NumberOfColumns = 20,
                NumberOfRows = worksheet.Rows.Count,
                Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
            });

            // Write DataTable content
            var sb = new StringBuilder();

            foreach (DataRow row in dataTable.Rows)
            {
                sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}", row[0], row[2], row[3], row[6], row[8], row[19]);

               
                
                var name = row[8];
                var tkt1 = row[0];
                var status = row[3];
                var priority = row[2];
                var ticketNumber = Convert.ToString(tkt1);
                var assignedTo = Convert.ToString(name);
                var trim = assignedTo.TrimEnd();
                var Status = Convert.ToString(status);
                var a = row[6];
                var TicketOpened = Convert.ToDateTime(a);

                Dictionary<string, string> Emial = new Dictionary<string, string>();
                //Emial.Add("Goutam", "xxxxx@hospital.com");
                ;
                


                foreach (KeyValuePair<string, string> entry in Emial)
                {
                    // do something with entry.Value or entry.Key

                    if (entry.Key != null)
                    {

                        if (trim == entry.Key)
                        {
                            var today = DateTime.Now;
                            //long hel = (long)today;
                            //get difference of two dates
                            var NumberOfDays = (today - TicketOpened).Days;

                            //Number of day calculation

                            if (NumberOfDays > 29)
                            {


                                try
                                {
                                    richTextBox1.AppendText("ASSIGENDTO:\t" + assignedTo);
                                    richTextBox1.AppendText("\t DATE:" + TicketOpened);

                                    // Email Part
                                    // Create the Outlook application.
                                    Outlook.Application oApp = new Outlook.Application();
                                    // Create a new mail item.
                                    Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                                    //add the body of the email
                                    oMsg.HTMLBody = "Hello " + assignedTo + ",<br/>You have been Tag to this ticket " + NumberOfDays +
                                        " days ago.Please complete this ticket as fast as possible.<br/> Ticket Number:" + ticketNumber +
                                        " <br/> Ticket Opend at: " + TicketOpened + "<br/>Priority:" + priority + "<br/>Staus:" + Status + "<br/><br/>Regards,<br/>abc";
                                    //Subject line
                                    oMsg.Subject = "Ticket Status";
                                    // Add a recipient.
                                    Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                                    // Change the recipient in the next line if necessary.
                                    Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(entry.Value);
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
                        }

                    }

                }
