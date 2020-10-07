using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;

namespace GFormsSyncData
{
    class Program
    {
        private static string FormName = "ITREBP_ITSupport_FaultLog";
        private static int TotalColumns = 10;
        private static int MaxRowId=2061;

        static void Main(string[] args)
        {
            Console.WriteLine("Google Forms Responses From Google Sheet");
            Console.WriteLine("========================================");
            Console.WriteLine("Form Name: " + FormName);
            Console.WriteLine("Total Columns: " + TotalColumns.ToString());
            Console.WriteLine("MaxRowsId: " + MaxRowId);

            Console.WriteLine("Fetching Data");
            DataSet dsGFormData = new DataSet("GFormData");
            dsGFormData.Tables.Add(GetGoogleSheetData(FormName, true, TotalColumns, MaxRowId));
            Console.WriteLine("Process Completed. Total Rows Fetched:"+dsGFormData.Tables[0].Rows.Count);

            string dsFileName = @"DS_" + DateTime.Now.Ticks.ToString() + ".xml";
            dsGFormData.WriteXml(dsFileName);
            Console.WriteLine("Data written to XML file: " + dsFileName);
            Console.ReadLine();
        }
        public static DataTable GetGoogleSheetData(string sheetFileName, bool exactMatch, int totalColumns, int maxRowId)
        {
            DataTable dt = new DataTable(sheetFileName);

            //Authenticate via OAuth2.0
            string keyFilePath = AppDomain.CurrentDomain.BaseDirectory + @"\GAccount-Key.p12";    // found in developer console
            string serviceAccountEmail = "641406182143-maf134v39q0r03s6k330a9lt6nm3tn2g@developer.gserviceaccount.com";   // found in developer console
            var certificate = new X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable);

            ServiceAccountCredential credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail) //create credential using certigicate
            {
                Scopes = new[] { "https://spreadsheets.google.com/feeds/" } //this scopr is for spreadsheets, check google scope FAQ for others
            }.FromCertificate(certificate));

            credential.RequestAccessTokenAsync(System.Threading.CancellationToken.None).Wait(); //request token

            var requestFactory = new GDataRequestFactory("My App User Agent");
            requestFactory.CustomHeaders.Add(string.Format("Authorization: Bearer {0}", credential.Token.AccessToken));

            SpreadsheetsService myService = new SpreadsheetsService("Spreadsheet-Photo-Tracking-App"); //create your old service
            myService.RequestFactory = requestFactory; //add new request factory to your old service

            //Get all Spreadsheets
            SpreadsheetQuery query = new SpreadsheetQuery();
            SpreadsheetFeed feed = myService.Query(query);

            //Console.WriteLine("Your spreadsheets:");
            foreach (SpreadsheetEntry entry in feed.Entries)
            {
                //Console.WriteLine(entry.Title.Text);

                //Get list of all worksheets
                if ((exactMatch && entry.Title.Text == sheetFileName) ||
                    (!exactMatch && entry.Title.Text.Contains(sheetFileName)))
                {
                    AtomLink link = entry.Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, null);

                    WorksheetQuery query2 = new WorksheetQuery(link.HRef.ToString());
                    WorksheetFeed wsFeed = myService.Query(query2);

                    int minRow = -1;
                    int maxRow = -1;
                    foreach (WorksheetEntry worksheet in wsFeed.Entries)
                    {
                        if (maxRowId > 0)
                        {
                            minRow = maxRowId;
                            maxRow = maxRowId;
                        }
                        else
                        {
                            minRow = 1;
                            maxRow = 1;
                        }

                        //Get Cell based Feed
                        AtomLink cellFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.CellRel, null);

                        CellQuery query3 = new CellQuery(cellFeedLink.HRef.ToString());
                        query3.MaximumColumn = (uint)totalColumns;

                        //Calculate Max Row available by fetching rows 1 by 1
                        query3.MinimumRow = (uint)minRow;
                        query3.MaximumRow = (uint)maxRow;
                        query3.ReturnEmpty = ReturnEmptyCells.yes;
                        CellFeed cFeed;
                        cFeed = myService.Query(query3);
                        while (((CellEntry)cFeed.Entries[0]).Value != null && (maxRow - maxRowId) < 200)
                        {
                            minRow++;
                            maxRow++;
                            query3.MinimumRow = (uint)minRow;
                            query3.MaximumRow = (uint)maxRow;
                            cFeed = myService.Query(query3);
                        }

                        //Fetch complete data
                        query3.MinimumRow = (uint)1;
                        query3.MaximumRow = (uint)maxRow;
                        cFeed = myService.Query(query3);

                        dt.DisplayExpression = maxRow.ToString();       //25Sep17: Additional line added in order to know the max row id of Google Sheet

                        int currRow = -1;
                        DataRow dr = null;
                        foreach (CellEntry curCell in cFeed.Entries)
                        {
                            if (curCell.Cell.Row != currRow || currRow == 1)
                            {
                                //New Row started
                                currRow = Convert.ToInt32(curCell.Cell.Row);

                                if (currRow == 1)
                                {
                                    //Its Header Row
                                    if (curCell.Cell.Value == "Timestamp")
                                        dt.Columns.Add("DateOfEntry");
                                    else
                                        dt.Columns.Add(curCell.Cell.Value
                                        .Replace(" ", "")
                                        .Replace(":", "")
                                        .Replace(".", "")
                                        .Replace("'", "")
                                        .Replace("&", "")
                                        .Replace("-", "")
                                        .Replace("/", "")
                                        .Replace("(", "")
                                        .Replace(")", ""));
                                    /*
                                    //EXCEL FORMULA to transform GSheets Column in DB Column Names
                                    =CONCATENATE("[",
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(
                                        SUBSTITUTE(A1," ","")
                                        ,":","")
                                        ,".","")
                                        ,"'","")
                                        ,"&","")
                                        ,"-","")
                                        ,"/","")
                                        ,"(","")
                                        ,")","")
                                        ,"]")
                                     */

                                    continue;
                                }
                                else
                                {
                                    if (dr != null && dr["DateOfEntry"].ToString().Trim().Length > 0)
                                        dt.Rows.Add(dr);
                                    dr = dt.NewRow();
                                }

                                Console.Write(".");
                            }
                            try
                            {
                                dr[Convert.ToInt32(curCell.Cell.Column) - 1] = ((dt.Columns[Convert.ToInt32(curCell.Cell.Column) - 1].ColumnName == "DateOfEntry") ? Convert.ToDateTime(curCell.Cell.Value).ToString("yyyy-MM-dd HH:mm:ss") : curCell.Cell.Value);
                            }
                            catch (Exception e)
                            {
                                if (dt.Columns[Convert.ToInt32(curCell.Cell.Column) - 1].ColumnName == "DateOfEntry")
                                {
                                    System.Diagnostics.EventLog.WriteEntry("SyncService", "DateTime.Now: " + DateTime.Now.ToString(), System.Diagnostics.EventLogEntryType.Information, 235);
                                    System.Diagnostics.EventLog.WriteEntry("SyncService", "Date Format: " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern, System.Diagnostics.EventLogEntryType.Information, 235);
                                    System.Diagnostics.EventLog.WriteEntry("SyncService", "Cell Value: " + curCell.Cell.Value, System.Diagnostics.EventLogEntryType.Information, 235);
                                }
                                throw e;
                            }
                        }
                        if (dr != null)
                            dt.Rows.Add(dr);
                        break;
                    }
                    break;
                }
            }

            Console.WriteLine();
            return dt;
        }
    }
}
