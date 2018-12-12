using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using M1_Data;
using System.Configuration;
using System.IO;
using ClosedXML.Excel;

namespace M1_PiXL
{
    public partial class Form1 : Form
    {
        public class ClientInfo
        {
            public string CompanyID = string.Empty;
            public string PiXLID = string.Empty;
            public string Email = string.Empty;
            public string OutputPathLegacy = string.Empty;
            public string OutputPathNew = string.Empty;

            public ClientInfo(DataRow row)
            {
                if (row.ItemArray.Count() != 5)
                    return;

                CompanyID = row[0].ToString();
                PiXLID = row[1].ToString();
                Email = row[2].ToString();
                OutputPathLegacy = row[3].ToString();
                OutputPathNew = row[4].ToString();
            }
        }

        private static M1_SQL M1SQL;
        private static M1_Email M1EMAIL;
        private static string PiXL_Match_Stored_Procedure = ConfigurationManager.AppSettings["PiXL_Match_Stored_Procedure"];
        private static string GetClients_Stored_Procedure = ConfigurationManager.AppSettings["GetClients_Stored_Procedure"];
        private static string EmailServer = ConfigurationManager.AppSettings["EmailServer"];
        private static string EmailPort = ConfigurationManager.AppSettings["EmailPort"];
        private static string EmailUser = ConfigurationManager.AppSettings["EmailUser"];
        private static string EmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
        private static string EmailSendCC = ConfigurationManager.AppSettings["EmailSendCC"];
        private static string EmailSendFrom = ConfigurationManager.AppSettings["EmailSendFrom"];
        private static string HTMLFile = @ConfigurationManager.AppSettings["HTMLFile"];
        private static int DayToProcess = int.Parse(@ConfigurationManager.AppSettings["DayToProcess"]);

        //public static int Rownumber = 1;
        //public static int Colnumber = 1;

        public Form1()
        {
            InitializeComponent();

            M1SQL = new M1_SQL(
            //ConfigurationManager.AppSettings["SQLUser"].ToString(),
            //ConfigurationManager.AppSettings["SQLPassword"].ToString(),
            "bgluckman", "Iruelakk2@",
            ConfigurationManager.AppSettings["SQLServer"].ToString(),
            ConfigurationManager.AppSettings["SQLDatabase"].ToString()
            );

            M1EMAIL = new M1_Email(EmailServer, EmailPort, EmailUser, EmailPassword);

            ProcessRecordsSP();
            Environment.Exit(1);
        }

        private void ProcessRecordsSP()
        {
            ClientInfo CurrentClientInfo;
            string outFolderLegacy = string.Empty;
            string outFolderNew = string.Empty;
            string PiXLAlias = string.Empty;
            string Line = string.Empty;
            XLWorkbook workbook;
            int RecordCount = 0;

            //XLWorkbook workbook = new XLWorkbook();
            //workbook.AddWorksheet("Matches");
            //bool firstpass = true;

            using (DataTable Clienttable = M1SQL.SQLGetDataFromStoredProcedure(GetClients_Stored_Procedure))
            {
                foreach (DataRow ClientRow in Clienttable.Rows)
                {
                    try
                    {
                        CurrentClientInfo = new ClientInfo(ClientRow);

                        using (DataTable PiXLResults = M1SQL.SQLGetDataFromStoredProcedure(PiXL_Match_Stored_Procedure, new List<M1_SQL.SQLParameterMap> {
                            new M1_SQL.SQLParameterMap("@CompanyID", CurrentClientInfo.CompanyID),
                            new M1_SQL.SQLParameterMap("@PiXLID", CurrentClientInfo.PiXLID),
                            new M1_SQL.SQLParameterMap("@DayToProcess", DayToProcess.ToString())}))
                        {
                            if (PiXLResults == null || PiXLResults.Rows.Count == 0)
                                continue;

                            PiXLAlias = CurrentClientInfo.OutputPathLegacy.Split('\\')[4].Split('_')[0];

                            outFolderLegacy = CurrentClientInfo.OutputPathLegacy + CurrentClientInfo.OutputPathLegacy.Split('\\')[4] + DateTime.Now.AddDays(-DayToProcess + 1).ToString("_yyyy_MM_dd") + ".xlsx";
                            outFolderNew = CurrentClientInfo.OutputPathNew + CurrentClientInfo.OutputPathNew.Split('\\')[4] + DateTime.Now.AddDays(-DayToProcess + 1).ToString("_yyyy_MM_dd") + ".xlsx";

                            workbook = new XLWorkbook();

                            //using (DataTable PiXLRecords = M1SQL.SQLGetDataFromStoredProcedure(GetPiXLRecords_Stored_Procedure, new List<M1_SQL.SQLParameterMap> {
                            //new M1_SQL.SQLParameterMap ("@Reseller", CurrentClientInfo.Reseller),
                            //new M1_SQL.SQLParameterMap ("@Client", CurrentClientInfo.Client),
                            //new M1_SQL.SQLParameterMap ("@Zip", CurrentClientInfo.Zip),
                            //new M1_SQL.SQLParameterMap ("@DaysToProcess", "1")}))
                            //    WritePiXLRecords(workbook, PiXLRecords, "PiXL Records");

                            //if (firstpass)
                            //{
                            //    foreach (DataColumn column in PiXLResults.Columns)
                            //    {
                            //        workbook.Worksheets.ToList()[0].Cell(Rownumber, Colnumber++).Value = column.ColumnName.ToLower().Contains("last seen") || column.ColumnName.ToLower().Contains("first seen") ? column.ColumnName + " Auto" : column.ColumnName;
                            //    }
                            //    firstpass = false;
                            //    Rownumber++;
                            //}

                            if (!Directory.Exists(CurrentClientInfo.OutputPathLegacy))
                                Directory.CreateDirectory(CurrentClientInfo.OutputPathLegacy);

                            if (!Directory.Exists(CurrentClientInfo.OutputPathNew))
                                Directory.CreateDirectory(CurrentClientInfo.OutputPathNew);

                            //WritePiXLRecordsCSV(PiXLResults, outFolderLegacy);
                            //WritePiXLRecordsCSV(PiXLResults, outFolderNew);
                            //RecordCount = WritePiXLRecordsXLSX(workbook, PiXLResults, "Matched Records");

                            //using (DataTable Unmatched = M1SQL.SQLGetDataFromStoredProcedure(GetUnmatched_Stored_Procedure, new List<M1_SQL.SQLParameterMap> {
                            //new M1_SQL.SQLParameterMap ("@Reseller", CurrentClientInfo.Reseller),
                            //new M1_SQL.SQLParameterMap ("@Client", CurrentClientInfo.Client),
                            //new M1_SQL.SQLParameterMap ("@Zip", CurrentClientInfo.Zip),
                            //new M1_SQL.SQLParameterMap ("@DaysToProcess", "1")}))
                            //    WritePiXLRecords(workbook, Unmatched, "UnMatched Records");

                            RecordCount = WritePiXLRecordsXLSX(workbook, PiXLResults, "records");
                            workbook.SaveAs(outFolderLegacy);
                            workbook.SaveAs(outFolderNew);

                            string EmailBody = PopulateBody(CurrentClientInfo.CompanyID, CurrentClientInfo.PiXLID, PiXLAlias, RecordCount.ToString());

                            //M1EMAIL.SendEmail(CurrentClientInfo.Email, new string[] { EmailSendCC }, EmailSendFrom, "Nightly Order Processed for PiXL: " + PiXLAlias, RecordCount.ToString() + " Records delivered. ", "");
                            //M1EMAIL.SendEmail("techteam@m1-data.com", new string[] { EmailSendCC }, EmailSendFrom, "Nightly Order Processed for PiXL: " + PiXLAlias, EmailBody, string.Empty, true);

                            //M1EMAIL.SendEmail(CurrentClientInfo.Email, new string[] { EmailSendCC }, EmailSendFrom, "Nightly Order Processed for PiXL: " + PiXLAlias, EmailBody, string.Empty, true);
                        }
                    }
                    catch (Exception exc)
                    {
                        new M1_Result(false, exc.ToString());
                    }
                }
            }
        }
        static private string PopulateBody(string CompanyID, string PiXLID, string PiXLAlias, string PiXLCount)
        {
            string body = string.Empty;
            using (StreamReader reader = new StreamReader(HTMLFile))
            {
                body = reader.ReadToEnd();
            }
            body = body.Replace("{CompanyID}", CompanyID);
            body = body.Replace("{PiXLID}", PiXLID);
            body = body.Replace("{PiXLAlias}", PiXLAlias);
            body = body.Replace("{PiXLCount}", PiXLCount);
            return body;
        }

        void WritePiXLRecordsCSV(DataTable table, string outFolder)
        {
            using (StreamWriter SW = new StreamWriter(outFolder))
            {
                string line = string.Empty;

                foreach (DataColumn Column in table.Columns)
                {
                    line += '\"' + Column.ColumnName + "\",";
                }

                SW.WriteLine(line.Substring(0, line.Length - 1));

                foreach(DataRow row in table.Rows)
                {
                    line = string.Empty;

                    foreach(var item in row.ItemArray)
                    {
                        line += '\"' + item.ToString() + "\",";
                    }
                    SW.WriteLine(line.Substring(0, line.Length - 1));
                }
            }
        }

        int WritePiXLRecordsXLSX(XLWorkbook workbook, DataTable table, string WorksheetName)
        {
            int Rownumber = 1;
            int Colnumber = 1;

            IXLWorksheet worksheet = workbook.AddWorksheet(WorksheetName);

            foreach (DataColumn column in table.Columns)
            {
                worksheet.Cell(Rownumber, Colnumber++).Value = column.ColumnName.ToLower().Contains("last seen") || column.ColumnName.ToLower().Contains("first seen") ? column.ColumnName + " Auto" : column.ColumnName;
                //workbook.Worksheets.ToList()[0].Cell(Rownumber, Colnumber++).Value = column.ColumnName.ToLower().Contains("last seen") || column.ColumnName.ToLower().Contains("first seen") ? column.ColumnName + " Auto" : column.ColumnName;
            }

            Rownumber++;

            foreach (DataRow row in table.Rows)
            {
                for (int c = 0; c < row.ItemArray.Length; c++)
                {
                    worksheet.Cell(Rownumber, c + 1).Value = row[c].ToString();
                    //workbook.Worksheets.ToList()[0].Cell(Rownumber, c + 1).Value = row[c].ToString();
                }

                Rownumber++;
            }
            return Rownumber;
        }
    }
}
