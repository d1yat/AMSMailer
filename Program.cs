using System.Text;
using System.Data;
using System.IO.Compression;
using OfficeOpenXml;
using MailKit.Net.Smtp;
using MimeKit;
using MySqlConnector;

namespace AMSMailer
{
    class Program
    {
        public static void Main(string[] args)
        {
            var today = DateTime.Now.DayOfWeek;
            string destdir = @"S:\Aplikasi\Project\AMS\AMSMailer\" + DateTime.Now.ToString("dd-MMM-yyyy");
            string olddir = @"S:\Aplikasi\Project\AMS\AMSMailer\" + DateTime.Now.AddDays(-2).ToString("dd-MMM-yyyy");
            string connectionStringSql = "Server=10.100.10.17;User ID=root;Password=@m5123;Database=ams_mrp_scms;Default Command Timeout=300;SSL Mode=None";

            if (today == DayOfWeek.Saturday && today == DayOfWeek.Sunday)
            {
                goto @end;
            }

            if (!Directory.Exists(destdir))
            {
                Directory.CreateDirectory(destdir);
            }

            if (Directory.Exists(olddir))
            {
                Directory.Delete(olddir, recursive: true);
            }

            try
            {
                DateTime currDate = DateTime.Today;
                DateTime prevDate = currDate.AddMonths(-1);
                DateTime date1 = new DateTime(prevDate.Year, prevDate.Month, 1);
                string path = string.Empty;
                string pathOld = string.Empty;
                DataTable? tbl = null;
                string content = string.Empty;
                string subject = string.Empty;
                string path2 = string.Empty;
                string pathOld2 = string.Empty;

                bool isMonday = today == DayOfWeek.Monday ? true : false;
                int end = isMonday ? 3 : 2;

                for (int i = 0; i <= end; i++)
                {
                    #region "History RN Cabang"

                    if (i == 0)
                    {
                        int rows = ExecuteStoredProcedure("sp_HistoryRNCabang", connectionStringSql);

                        tbl = new DataTable();
                        tbl = ExecuteSql("select * from LG_tempHistoryRNCabang_Jobs", connectionStringSql);
                        Console.WriteLine("Save to excel");

                        path = destdir + "\\History-RN-Cabang-" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        pathOld = destdir + "\\History-RN-Cabang-" + DateTime.Now.AddDays(-1).ToString("ddMMyyyy") + ".xlsx";

                        using (ExcelPackage ExcelPkg = new ExcelPackage())
                        {
                            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("DOPHarmanet");
                            wsSheet1.Cells["A1"].LoadFromDataTable(tbl, true);
                            ExcelPkg.SaveAs(new FileInfo(path));

                        }

                        content = "Terlampir di attachment data History RN Cabang. Periode " + date1.ToString("dd MMM yyyy") + " s/d " + currDate.ToString("dd MMM yyyy") + ".";
                        subject = "History RN Cabang";
                    }

                    #endregion

                    #region "History PO Pharmanet"

                    else if (i == 1)
                    {
                        int rows = ExecuteStoredProcedure("sp_HistoryPOPharmanet", connectionStringSql);

                        tbl = new DataTable();
                        tbl = ExecuteSql("select * from Temp_Proses_PO_Pharmanet", connectionStringSql);
                        Console.WriteLine("Save to excel");

                        path = destdir + "\\History-PO-Pharmanet-" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        pathOld = destdir + "\\History-PO-Pharmanet-" + DateTime.Now.AddDays(-1).ToString("ddMMyyyy") + ".xlsx";

                        using (ExcelPackage ExcelPkg = new ExcelPackage())
                        {
                            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("PO Pharmanet");
                            wsSheet1.Cells["A1"].LoadFromDataTable(tbl, true);
                            ExcelPkg.SaveAs(new FileInfo(path));

                        }

                        content = "Terlampir di attachment data History PO Pharmanet. Periode " + date1.ToString("dd MMM yyyy") + " s/d " + currDate.ToString("dd MMM yyyy") + ".";
                        subject = "History PO Pharmanet";
                    }

                    #endregion

                    #region "Laporan PL yang Diterima"

                    else if (i == 2)
                    {
                        int rows = ExecuteStoredProcedure("sp_LaporanPLDiterima", connectionStringSql);

                        tbl = new DataTable();
                        tbl = ExecuteSql("select * from ProsesPharmanet", connectionStringSql);
                        Console.WriteLine("Save to excel");

                        path = destdir + "\\Laporan-PL-yang-Diterima-" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        pathOld = destdir + "\\Laporan-PL-yang-Diterima-" + DateTime.Now.AddDays(-1).ToString("ddMMyyyy") + ".xlsx";

                        using (ExcelPackage ExcelPkg = new ExcelPackage())
                        {
                            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Per PL");
                            wsSheet1.Cells["A1"].LoadFromDataTable(tbl, true);
                            ExcelPkg.SaveAs(new FileInfo(path));

                        }

                        tbl.Clear();

                        tbl = new DataTable();
                        tbl = ExecuteSql("select * from datapharmanet where itemterdaftar <> 'item terdaftar' or statusharga <> 'harga sama' or ketqty <> 'qty sama' or ketbatch <> 'batch sama'", connectionStringSql);
                        Console.WriteLine("Save to excel");

                        path2 = destdir + "\\Laporan-PL-detail-yang-Diterima-" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        pathOld2 = destdir + "\\Laporan-PL-detail-yang-Diterima-" + DateTime.Now.AddDays(-1).ToString("ddMMyyyy") + ".xlsx";

                        using (ExcelPackage ExcelPkg = new ExcelPackage())
                        {
                            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Detail Item");
                            wsSheet1.Cells["A1"].LoadFromDataTable(tbl, true);
                            ExcelPkg.SaveAs(new FileInfo(path2));


                        }

                        content = "Terlampir di attachment data Laporan PL yang Diterima. Periode " + date1.ToString("dd MMM yyyy") + " s/d " + currDate.ToString("dd MMM yyyy") + ".";
                        subject = "Laporan PL yang Diterima";
                    }

                    #endregion

                    #region "Laporan Data Pharmanet"

                    else if (i == 3)
                    {
                        int rows = ExecuteStoredProcedure("sp_LaporanDataPharmanet", connectionStringSql);

                        tbl = new DataTable();
                        tbl = ExecuteSql("select * from Temp_Porses_PLPharmanet", connectionStringSql);
                        Console.WriteLine("Save to excel");

                        path = destdir + "\\Laporan-Data-Pharmanet-" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                        pathOld = destdir + "\\Laporan-Data-Pharmanet-" + DateTime.Now.AddDays(-1).ToString("ddMMyyyy") + ".xlsx";

                        using (ExcelPackage ExcelPkg = new ExcelPackage())
                        {
                            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("PL Pharmanet");
                            wsSheet1.Cells["A1"].LoadFromDataTable(tbl, true);
                            ExcelPkg.SaveAs(new FileInfo(path));

                        }

                        content = "Terlampir di attachment data Laporan Data Pharmanet. Periode " + date1.ToString("dd MMM yyyy") + " s/d " + currDate.ToString("dd MMM yyyy") + ".";
                        subject = "Laporan Data Pharmanet";
                    }

                    #endregion

                    tbl?.Clear();
                    tbl?.Dispose();

                    #region Send Mail

                    if (File.Exists(path))
                    {
                        Console.Write($"Send Email... {subject}. ");

                        var message = new MimeMessage();
                        message.From.Add(new MailboxAddress("Supply Chain Management System", "scms.dophar@ams.co.id"));
                        message.Subject = subject;

                        if (i <= 3)
                        {
                            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                            // Development
                            string recipient = Path.Combine(baseDirectory, @"..\..\..\recipient.txt");

                            if (!File.Exists(recipient))
                            {
                                // Production
                                recipient = Path.Combine(baseDirectory, @".\recipient.txt");

                                if (!File.Exists(recipient))
                                {
                                    throw new Exception(string.Concat(recipient, " does not exists."));
                                }
                            }

                            string[] addresses = File.ReadAllLines(recipient).Where(address => !address.StartsWith("'")).ToArray();

                            foreach (var address in addresses)
                            {
                                message.To.Add(new MailboxAddress("", address));
                            }
                        }

                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine(content);
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.AppendLine("Terima Kasih,");
                        sb.AppendLine("AMS - MIS Team");

                        var bodyBuilder = new BodyBuilder()
                        {
                            TextBody = sb.ToString()
                        };

                        bodyBuilder.Attachments.Add(path);
                        if (i == 2 && File.Exists(path2))
                        {
                            bodyBuilder.Attachments.Add(path2);
                        }

                        message.Body = bodyBuilder.ToMessageBody();

                        sb.Length = 0;

                        try
                        {
                            using (var mail = new SmtpClient())
                            {
                                if (i <= 3)
                                {
                                    mail.Connect("10.100.10.9", 25);
                                    mail.Authenticate("scms.dophar@ams.co.id", "scmsdophar");
                                    mail.Send(message);
                                    mail.Disconnect(true);

                                    Console.WriteLine("\t[Done]");
                                }
                            }

                            if (File.Exists(pathOld))
                            {
                                File.Delete(pathOld);
                            }
                        }
                        catch (Exception ex)
                        {
                            string[] line = { ex.Message, "", "" };
                            System.IO.File.WriteAllLines(@"D:\AMSMailer\error-send-email.txt", line);
                            Console.WriteLine(ex.Message);
                            Console.WriteLine(ex.StackTrace);
                        }
                    }

                    #endregion
                }

            }
            catch (Exception ex)
            {
                string[] line = { ex.Message, "", "" };
                System.IO.File.WriteAllLines(@"D:\AMSMailer\error.txt", line);
                Console.WriteLine(ex.Message);
            }

        @end:
            Console.WriteLine("Email tidak dikirimkan pada hari Sabtu & Minggu.");
        }

        public static DataTable ExecuteSql(string query, string connectionString)
        {
            using (MySqlConnection conSql = new MySqlConnection(connectionString))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                DataTable tbl = new DataTable();
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conSql);
                adapter.Fill(tbl);
                return tbl;
            }
        }

        public static int ExecuteStoredProcedure(string storedProcedure, string connectionString)
        {
            using (MySqlConnection conSql = new MySqlConnection(connectionString))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandText = storedProcedure;
                cmd.Connection = conSql;
                cmd.CommandType = CommandType.StoredProcedure;
                return cmd.ExecuteNonQuery();
            }
        }

        public static string Compress(string pathToFile)
        {
            FileInfo fileToBeGZipped = new FileInfo(pathToFile);
            FileInfo gzipFileName = new FileInfo(string.Concat(fileToBeGZipped.FullName, ".gz"));

            using (FileStream fileToBeZippedAsStream = fileToBeGZipped.OpenRead())
            {
                using (FileStream gzipTargetAsStream = gzipFileName.Create())
                {
                    using (GZipStream gzipStream = new GZipStream(gzipTargetAsStream, CompressionMode.Compress))
                    {
                        try
                        {
                            byte[] b = new byte[fileToBeZippedAsStream.Length];
                            int read = fileToBeZippedAsStream.Read(b, 0, b.Length);
                            while (read > 0)
                            {
                                gzipStream.Write(b, 0, read);
                                read = fileToBeZippedAsStream.Read(b, 0, b.Length);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }

            return gzipFileName.FullName;
        }
    }
}
