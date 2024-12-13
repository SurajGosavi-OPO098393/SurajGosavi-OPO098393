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
using System.Linq.Expressions;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ICICI_Trail_Upload_File_Service
{


    public partial class Service1 : ServiceBase
    {
        string logFilePath = $@"D:\ICICI_Trail_ExecutionLog\Log\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            System.Timers.Timer myTimer = new System.Timers.Timer();
            myTimer.Elapsed += new ElapsedEventHandler(timer1_Tick);
            myTimer.Interval = Convert.ToInt64(ConfigurationSettings.AppSettings["Timer"].ToString());
            myTimer.Enabled = true;
            myTimer.Start();
        }

        protected override void OnStop()
        {

        }

        public void OnDebug()
        {


            System.Timers.Timer myTimer = new System.Timers.Timer();
            myTimer.Elapsed += new ElapsedEventHandler(timer1_Tick);
            myTimer.Interval = Convert.ToInt64(ConfigurationSettings.AppSettings["Timer"].ToString());
            myTimer.Enabled = true;
            myTimer.Start();


        }


        private void ExecuteAndSaveReport_ICICI_GGN_Crds()
        {
            
                        try
                        {
                            LogMessage(logFilePath, "creating path ICICI_GGN_Crds...");
                            string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Change this to your desired output path
                            //string outputDirectory1 = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Local Directory

                            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));


                            LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                            LogMessage(logFilePath, "path Created");
                            // string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt"; // Date-specific log file path
                            LogMessage(logFilePath, "ICICI_GGN_Crds Execution started.");

                        }
                        catch (Exception ex)
                        {
                            //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                            LogMessage(logFilePath, $"Error occurred: {ex.Message}");
                        }
                        try
                        {
                            LogMessage(logFilePath, "getdata TRAIL_UPLOAD_FILE_GGN_Automation");
                            DataTable dt = new DataTable();
                            string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CRD"].ToString();
                            using (SqlConnection con = new SqlConnection(connectionString))
                            {
                                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_Automation", con);
                                cmd.Parameters.AddWithValue("@Operation", "Card_Data");
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandTimeout = 0;

                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                da.Fill(dt);
                            }

                            int rowCount = dt.Rows.Count;
                            LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                            if (rowCount > 0)
                            {
                            LogMessage(logFilePath, "Inside get chunk file");
                            int fileIndex = 1;
                                int chunkSize = 9999;
                                string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Change this to your desired output path
                                //string outputDirectory = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Local Directory

                                string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                                Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                                LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                                Directory.CreateDirectory(outputDirectory); // Ensure the directory exists
                                LogMessage(logFilePath, "Directory Created");
                                for (int i = 0; i < rowCount; i += chunkSize)
                                {
                                    DataTable chunkTable = dt.AsEnumerable()
                                                             .Skip(i)
                                                             .Take(chunkSize)
                                                             .CopyToDataTable();

                                    //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                                    string fileName = Path.Combine(datewiseFolder, $"Card_{fileIndex}.csv");
                                    LogMessage(logFilePath, "before ExportChunkToExcel");
                                    ExportChunkToExcel(chunkTable, fileName);
                                    LogMessage(logFilePath, "AFTER ExportChunkToExcel");
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


        private void ExecuteAndSaveReport_ICICI_GGN_Crds_Manual()
        {

            try
            {
                LogMessage(logFilePath, "creating path ICICI_GGN_Crds_Manual...");
                string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Change this to your desired output path
                //string outputDirectory1 = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Local Directory

                string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));


                LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                LogMessage(logFilePath, "path Created");
                // string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt"; // Date-specific log file path
                LogMessage(logFilePath, "ICICI_GGN_Crds_Manual Execution started.");

            }
            catch (Exception ex)
            {
                //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
            try
            {
                LogMessage(logFilePath, "getdata TRAIL_UPLOAD_FILE_GGN_Automation");
                DataTable dt = new DataTable();
                string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CRD"].ToString();
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_Automation", con);
                    cmd.Parameters.AddWithValue("@Operation", "Card_Manual_Data");
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                }

                int rowCount = dt.Rows.Count;
                LogMessage(logFilePath, $"Data fetched: {rowCount} rows.");
                if (rowCount > 0)
                {
                    LogMessage(logFilePath, "Inside get chunk file");
                    int fileIndex = 1;
                    int chunkSize = 9999;
                    string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Change this to your desired output path
                    //string outputDirectory = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CARDS"; // Local Directory

                    string datewiseFolder = Path.Combine(outputDirectory, DateTime.Now.ToString("yyyy-MM-dd"));
                    Directory.CreateDirectory(datewiseFolder); // Ensure the folder exists

                    LogMessage(logFilePath, $"Output directory created: {datewiseFolder}");

                    Directory.CreateDirectory(outputDirectory); // Ensure the directory exists
                    LogMessage(logFilePath, "Directory Created");
                    for (int i = 0; i < rowCount; i += chunkSize)
                    {
                        DataTable chunkTable = dt.AsEnumerable()
                                                 .Skip(i)
                                                 .Take(chunkSize)
                                                 .CopyToDataTable();

                        //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                        string fileName = Path.Combine(datewiseFolder, $"Card_Manual_{fileIndex}.csv");
                        LogMessage(logFilePath, "before ExportChunkToExcel");
                        ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, "AFTER ExportChunkToExcel");
                        LogMessage(logFilePath, $"Manual File generated: {fileName}");
                        fileIndex++;
                    }

                    LogMessage(logFilePath, "Manual Report generation completed successfully.");

                    Console.WriteLine($"Manaul Reports generated successfully in {outputDirectory}");
                }
                else
                {
                    LogMessage(logFilePath, "No Manual data to export.");
                    Console.WriteLine("No Manual data to export.");
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
            
                        try
                        {
                            LogMessage(logFilePath, "creating path ICICI_GGN_CFL...");
                            string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Change this to your desired output path
                            //string outputDirectory1 = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Local Directory

                            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists


                            LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                            LogMessage(logFilePath, "path Created");
                            LogMessage(logFilePath, "ICICI_GGN_CFL Execution started.");

                        }
                        catch (Exception ex)
                        {
                            //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                            LogMessage(logFilePath, $"Error occurred: {ex.Message}");
                        }

                        try
                        {

                            DataTable dt = new DataTable();
                            string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CFL"].ToString();
                            using (SqlConnection con = new SqlConnection(connectionString))
                            {
                                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_CFL_Automation", con);
                                cmd.Parameters.AddWithValue("@Operation", "CFL_Data");
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
                                string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Change this to your desired output path
                                //string outputDirectory = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Local Directory
                                
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

                                    //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CFL_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                                    string fileName = Path.Combine(datewiseFolder, $"Loan_{fileIndex}.csv");
                                    ExportChunkToExcel(chunkTable, fileName);
                                    LogMessage(logFilePath, $"File generated: {fileName}");
                                    fileIndex++;
                                }
                                LogMessage(logFilePath, "Report generation completed successfully.");

                                LogMessage(logFilePath, $"Reports generated successfully in {outputDirectory}");
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

        private void ExecuteAndSaveReport_ICICI_GGN_CFL_Manual()
        {

            try
            {
                LogMessage(logFilePath, "creating path ICICI_GGN_CFL Manual...");
                string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Change this to your desired output path
                //string outputDirectory1 = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Local Directory

                string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists


                LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                LogMessage(logFilePath, "path Created");
                LogMessage(logFilePath, "ICICI_GGN_CFL Manual Execution started.");

            }
            catch (Exception ex)
            {
                //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }

            try
            {

                DataTable dt = new DataTable();
                string connectionString = ConfigurationSettings.AppSettings["Connection_GGN_CFL"].ToString();
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_GGN_CFL_Automation", con);
                    cmd.Parameters.AddWithValue("@Operation", "CFL_Manual_Data");
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
                    string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Change this to your desired output path
                    //string outputDirectory = @"C:\\ICICI_TrailFile_GGN\ICICI_GGN_CFL"; // Local Directory

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

                        //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_GGN_CFL_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                        string fileName = Path.Combine(datewiseFolder, $"Loan_Manual_{fileIndex}.csv");
                        ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"Manual File generated: {fileName}");
                        fileIndex++;
                    }
                    LogMessage(logFilePath, "Manual Report generation completed successfully.");

                    LogMessage(logFilePath, $"Manual Reports generated successfully in {outputDirectory}");
                }
                else
                {
                    LogMessage(logFilePath, "No Manual data to export.");
                    Console.WriteLine("No Manual data to export.");
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
            
                        try
                        {
                            LogMessage(logFilePath, "creating path ICICI_BLR_CRD...");
                            string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Change this to your desired output path
                            //string outputDirectory1 = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Local Directory

                            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

                            LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                            LogMessage(logFilePath, "path Created");
                            LogMessage(logFilePath, "ICICI_BLR_CARDS Execution started.");

                        }
                        catch (Exception ex)
                        {
                            //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                            LogMessage(logFilePath, $"Error occurred: {ex.Message}");
                        }
                        LogMessage(logFilePath, "ICICI_BLR_CRD Execution started.");

                        try
                        {
                            DataTable dt = new DataTable();
                            string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_CRD"].ToString();
                            using (SqlConnection con = new SqlConnection(connectionString))
                            {
                                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_Automation", con);
                                cmd.Parameters.AddWithValue("@Operation", "Card_Data");
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
                                string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Change this to your desired output path
                                //string outputDirectory = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Local Directory

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

                                    //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                                    string fileName = Path.Combine(datewiseFolder, $"Card_{fileIndex}.csv");
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

        private void ExecuteAndSaveReport_ICICI_BLR_CRD_Manual()
        {

            try
            {
                LogMessage(logFilePath, "creating path ICICI_BLR_CRD Manual...");
                string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Change this to your desired output path
                //string outputDirectory1 = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Local Directory

                string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists

                LogMessage(logFilePath, "folder path: " + datewiseFolder1);

                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                LogMessage(logFilePath, "path Created");
                LogMessage(logFilePath, "ICICI_BLR_CARDS Manual Execution started.");

            }
            catch (Exception ex)
            {
                //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
            LogMessage(logFilePath, "ICICI_BLR_CRD Manual Execution started.");

            try
            {
                DataTable dt = new DataTable();
                string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_CRD"].ToString();
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_Automation", con);
                    cmd.Parameters.AddWithValue("@Operation", "Card_Manual_Data");
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
                    string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Change this to your desired output path
                    //string outputDirectory = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_CARDS"; // Local Directory

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

                        //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_CARD_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                        string fileName = Path.Combine(datewiseFolder, $"Card_Manual_{fileIndex}.csv");
                        ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                    }

                    LogMessage(logFilePath, "Manual Report generation completed successfully.");
                    Console.WriteLine($"Manual Reports generated successfully in {outputDirectory}");
                }
                else
                {
                    LogMessage(logFilePath, "No Manual data to export.");
                    Console.WriteLine("No Manual data to export.");
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
            
                        try
                        {
                            LogMessage(logFilePath, "creating path ICICI_BLR_PL...");
                            string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Change this to your desired output path
                            //string outputDirectory1 = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Local Directory

                            string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists


                            Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                            LogMessage(logFilePath, "path Created");
                            LogMessage(logFilePath, "ICICI_BLR_PL Execution started.");
                        }
                        catch (Exception ex)
                        {
                            //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                            LogMessage(logFilePath, $"Error occurred: {ex.Message}");
                        }
                        try
                        {
                            DataTable dt = new DataTable();
                            string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_PL"].ToString();
                            using (SqlConnection con = new SqlConnection(connectionString))
                            {
                                SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_PL_Automation", con);
                                cmd.Parameters.AddWithValue("@Operation", "PL_Data");
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
                                string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Change this to your desired output path
                                //string outputDirectory = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Local Directory

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

                                    //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                                    string fileName = Path.Combine(datewiseFolder, $"Loan_{fileIndex}.csv");
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



        private void ExecuteAndSaveReport_ICICI_BLR_PL_Manual()
        {

            try
            {
                LogMessage(logFilePath, "creating path ICICI_BLR_PL Manual...");
                string outputDirectory1 = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Change this to your desired output path
                //string outputDirectory1 = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Local Directory

                string datewiseFolder1 = Path.Combine(outputDirectory1, DateTime.Now.ToString("yyyy-MM-dd"));
                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists


                Directory.CreateDirectory(datewiseFolder1); // Ensure the folder exists
                LogMessage(logFilePath, "path Created");
                LogMessage(logFilePath, "ICICI_BLR_PL Manual Execution started.");
            }
            catch (Exception ex)
            {
                //string logFilePath = $@"D:\ExecutionLog_{DateTime.Now:yyyy-MM-dd}.txt";
                LogMessage(logFilePath, $"Error occurred: {ex.Message}");
            }
            try
            {
                DataTable dt = new DataTable();
                string connectionString = ConfigurationSettings.AppSettings["Connection_BLR_PL"].ToString();
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand("USP_TRAIL_UPLOAD_FILE_BLR_PL_Automation", con);
                    cmd.Parameters.AddWithValue("@Operation", "PL_Data_Manual");
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
                    string outputDirectory = @"\\172.24.11.42\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Change this to your desired output path
                    //string outputDirectory = @"C:\\ICICI_TrailFile_BLR\ICICI_BLR_PL"; // Local Directory

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

                        //string fileName = Path.Combine(datewiseFolder, $"TRAIL_UPLOAD_FILE_BLR_PL_Report_Part_{DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss")}_{fileIndex}.csv");
                        string fileName = Path.Combine(datewiseFolder, $"Loan_Manual_{fileIndex}.csv");
                        ExportChunkToExcel(chunkTable, fileName);
                        LogMessage(logFilePath, $"File generated: {fileName}");
                        fileIndex++;
                    }
                    LogMessage(logFilePath, "Manual Report generation completed successfully.");
                    Console.WriteLine($"Manual Reports generated successfully in {outputDirectory}");
                }
                else
                {
                    LogMessage(logFilePath, "No Manual data to export.");
                    Console.WriteLine("No Manual data to export.");
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
            // Log entry with timestamp
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";

            // Append log entry to the specified log file
            File.AppendAllText(logFilePath, logEntry);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            

            if (DateTime.Now.TimeOfDay.Hours == 19 && DateTime.Now.TimeOfDay.Minutes == 00) // 8:00 PM
            {


                LogMessage(logFilePath, "Execution Started");


                ExecuteAndSaveReport_ICICI_GGN_Crds();
                LogMessage(logFilePath, "ICICI_GGN_Crds data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_GGN_Crds_Manual();
                LogMessage(logFilePath, "ICICI_GGN_Crds_Manual data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_GGN_CFL();
                LogMessage(logFilePath, "ICICI_GGN_CFL data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_GGN_CFL_Manual();
                LogMessage(logFilePath, "ICICI_GGN_CFL_Manual data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_BLR_CRD();
                LogMessage(logFilePath, "ICICI_BLR_CRD data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_BLR_CRD_Manual();
                LogMessage(logFilePath, "ICICI_BLR_CRD_Manual data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_BLR_PL();
                LogMessage(logFilePath, "ICICI_BLR_PL data generated");
                LogMessage(logFilePath, "");

                ExecuteAndSaveReport_ICICI_BLR_PL_Manual();
                LogMessage(logFilePath, "ICICI_BLR_PL_Manual data generated");
                LogMessage(logFilePath, "");
            }
            else
            {
                LogMessage(logFilePath, "Service is Running...");
            }
        }
    }
}
