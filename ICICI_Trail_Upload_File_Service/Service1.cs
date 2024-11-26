using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ICICI_Trail_Upload_File_Service
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }

        public void OnDebug()
        {
            //ExecuteAndSaveReport();

            if (DateTime.Now.TimeOfDay.Hours == 20 /*&& DateTime.Now.TimeOfDay.Minutes == 0*/) // 8:00 PM
            {
                ExecuteAndSaveReport_ICICI_GGN_Crds();
                ExecuteAndSaveReport_ICICI_GGN_CFL();
                ExecuteAndSaveReport_ICICI_BLR_CRD();
                ExecuteAndSaveReport_ICICI_BLR_PL();
            }

        }


        private void ExecuteAndSaveReport_ICICI_GGN_Crds()
        {

            string outputDirectory1 = @"C:\Reports\ICICI_GGN_CARDS"; // Change this to your desired output path

            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

            string logFilePath = @"C:\Reports\ICICI_GGN_CARDS\ExecutionLog.txt"; // Log file path
            LogMessage(logFilePath, "Execution started.");

            try
            {

                DataTable dt = new DataTable();
                string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CRD"].ToString();
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_Automation", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                }

                int rowCount = dt.Rows.Count;
                LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                if (rowCount > 0)
                {
                    int fileIndex = 1;
                    int chunkSize = 9999;
                    string outputDirectory = @"C:\Reports\ICICI_GGN_CARDS"; // Change this to your desired output path

                    string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                    Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                    LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                    Directory.CreateDirectory(outputDirectory); // Ensure the directory exists

                    for (int i = 0; i < rowCount; i += chunkSize)
                    {
                        DataTable chunkTable = dt.AsEnumerable()
                                                 .Skip(i)
                                                 .Take(chunkSize)
                                                 .CopyToDataTable();

                        string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.xlsx");
                        ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                    }

                    LogMessage(logFilePath, "Report generation completed successfully.");

                    Console.WriteLine($"Reports generated successfully in {outputDirectory}");
                }
                else
                {
                    LogMessage(logFilePath, "No data to export.");
                    Console.WriteLine("No data to export.");
                }
            }
            catch (Exception ex)
            {
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
        }



        //------------------------------------------------------------------------------------------------


        private void ExecuteAndSaveReport_ICICI_GGN_CFL()
        {

            string outputDirectory1 = @"C:\Reports\ICICI_GGN_CFL"; // Change this to your desired output path

            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

            string logFilePath = @"C:\Reports\ICICI_GGN_CFL\ExecutionLog.txt"; // Log file path
            LogMessage(logFilePath, "Execution started.");

            try
            {

            DataTable dt = new DataTable();
            string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CFL"].ToString();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_CFL_Automation", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }

            int rowCount = dt.Rows.Count;
                LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                if (rowCount > 0)
            {
                int fileIndex = 1;
                int chunkSize = 9999;
                string outputDirectory = @"C:\Reports\ICICI_GGN_CFL"; // Change this to your desired output path

                string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                    LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                    Directory.CreateDirectory(outputDirectory); // Ensure the directory exists

                for (int i = 0; i < rowCount; i += chunkSize)
                {
                    DataTable chunkTable = dt.AsEnumerable()
                                             .Skip(i)
                                             .Take(chunkSize)
                                             .CopyToDataTable();

                    string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CFL_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.xlsx");
                    ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                }
                    LogMessage(logFilePath, "Report generation completed successfully.");

                    Console.WriteLine($"Reports generated successfully in {outputDirectory}");
            }
                else
                {
                    LogMessage(logFilePath, "No data to export.");
                    Console.WriteLine("No data to export.");
                }
            }
            catch (Exception ex)
            {
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
        }


        //------------------------------------------------------------------------------------------------


        private void ExecuteAndSaveReport_ICICI_BLR_CRD()
        {

            string outputDirectory1 = @"C:\Reports\ICICI_BLR_CARDS"; // Change this to your desired output path

            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

            string logFilePath = @"C:\Reports\ICICI_BLR_CARDS\ExecutionLog.txt"; // Log file path
            LogMessage(logFilePath, "Execution started.");

            try
            {
                DataTable dt = new DataTable();
            string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_CRD"].ToString();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_Automation", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }

            int rowCount = dt.Rows.Count;
                LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                if (rowCount > 0)
            {
                int fileIndex = 1;
                int chunkSize = 9999;
                string outputDirectory = @"C:\Reports\ICICI_BLR_CARDS"; // Change this to your desired output path

                string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                    LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                    Directory.CreateDirectory(outputDirectory); // Ensure the directory exists

                for (int i = 0; i < rowCount; i += chunkSize)
                {
                    DataTable chunkTable = dt.AsEnumerable()
                                             .Skip(i)
                                             .Take(chunkSize)
                                             .CopyToDataTable();

                    string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.xlsx");
                    ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                }

                    LogMessage(logFilePath, "Report generation completed successfully.");
                    Console.WriteLine($"Reports generated successfully in {outputDirectory}");
            }
                else
                {
                    LogMessage(logFilePath, "No data to export.");
                    Console.WriteLine("No data to export.");
                }
            }
            catch (Exception ex)
            {
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
        }



        //------------------------------------------------------------------------------------------------


        private void ExecuteAndSaveReport_ICICI_BLR_PL()
        {

            string outputDirectory1 = @"C:\Reports\ICICI_BLR_PL"; // Change this to your desired output path

            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

            string logFilePath = @"C:\Reports\ICICI_BLR_PL\ExecutionLog.txt"; // Log file path
            LogMessage(logFilePath, "Execution started.");

            try
            {
                DataTable dt = new DataTable();
            string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_PL"].ToString();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_PL_Automation", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }

            int rowCount = dt.Rows.Count;
                LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                if (rowCount > 0)
            {
                int fileIndex = 1;
                int chunkSize = 9999;
                string outputDirectory = @"C:\Reports\ICICI_BLR_PL"; // Change this to your desired output path

                string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                    LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                    Directory.CreateDirectory(outputDirectory); // Ensure the directory exists

                for (int i = 0; i < rowCount; i += chunkSize)
                {
                    DataTable chunkTable = dt.AsEnumerable()
                                             .Skip(i)
                                             .Take(chunkSize)
                                             .CopyToDataTable();

                    string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.xlsx");
                    ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                }
                    LogMessage(logFilePath, "Report generation completed successfully.");
                    Console.WriteLine($"Reports generated successfully in {outputDirectory}");
            }
            else
            {
                LogMessage(logFilePath, "No data to export.");
                Console.WriteLine("No data to export.");
            }
        }
            catch (Exception ex)
            {
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
    }
}

        private static void ExportChunkToExcel(DataTable data, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromDataTable(data, true); // Load data with headers

                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }

        private void LogMessage(string logFilePath, string message)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";
            File.AppendAllText(logFilePath, logEntry);
        }


    }
}
