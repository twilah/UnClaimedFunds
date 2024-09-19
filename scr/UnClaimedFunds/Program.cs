using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

///<summary>
///This module is used to consolidate all the files in the Unclaimed Funds folder.  
///It is designed to process unclaimed funds data from various sources, consolidate it, and log the process. 
///The application handles both CSV and Excel files, organizing them based on specific criteria
///Key Functionalities
///•	The application is designed to process data from a structured directory of unclaimed funds, specifically targeting files from the year 2013 that do not belong to the "RPS" category.
///•	It supports processing both CSV and Excel files, consolidating the data into output CSV files organized by year and type (e.g., Pension, Group, Survivor).
///•	The process includes logging actions and handling file access conflicts, which is crucial for dealing with files that might be in use by other processes.
///•	The application uses external libraries like CsvHelper and ExcelDataReader for reading and writing CSV and Excel files, respectively.
///</summary>
///<para> Module   :  Program.cs                                                                                  </para>
///<para> Author   :  Twila Williams                                                                              </para>
///<para> Date     :  06/06/2024                                                                                  </para>
///<para>                                                                                                         </para>
///<para> Change Date  Developer Name          Change Description                                                 </para>
///<para> -----------  ----------------------  ------------------------------------------------------------------ </para>
///<para> 06/06/2024   Twila Williams          Created original version                                           </para>

namespace UnClaimedFunds
{
    static class Program
    {
        #region Variables
        private static string _unclaimdedFundsFolder = @"\\DPT42-2DKVDW2\Users\duke\Documents\UnclaimedFunds\ALL STATES Unclaimed Property Audit - DMF";
        private static string _outputFolder = @"\\DPT42-2DKVDW2\Users\duke\Documents\UnclaimedFunds\CONSOLIDATED DMF";
        private static string _logFolder = @"\\DPT42-2DKVDW2\Users\duke\Documents\UnclaimedFunds\CONSOLIDATED DMF\Logs";
        private static StreamWriter _logger;
        private static string _outputFile = string.Empty;
        private static string _folderYear = string.Empty;
        #endregion

        /// <summary>
        /// The main entry point for the application. It creates the necessary directories, initializes the log file, 
        /// processes all folders in the unclaimed funds directory and handles any exceptions that may occur.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                CreateDirectory();
                CreateLogFile(_logFolder);
                ProcessAllFoldersInDirectory(_unclaimdedFundsFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                _logger.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                _logger.Flush();
                _logger.Close();
            }

        }

        /// <summary>
        /// Checks if the output and log directories exist, and creates them if they don't.
        /// </summary>
          public static void CreateDirectory()
        {
            if (!Directory.Exists(_outputFolder))
            {
                Directory.CreateDirectory(_outputFolder);
            }
            if (!Directory.Exists(_logFolder))
            {
                Directory.CreateDirectory(_logFolder);
            }
        }

        /// <summary>
        /// Initializes the log file in the specified log folder with a timestamped filename.
        /// <paramref name="logFolder"/>The directory path where the log file will be created.
        public static void CreateLogFile(string logFolder)
        {
            string logFileName = string.Format(CultureInfo.CurrentCulture, "Log-{0:yyyy-MM-dd_H-mm-ss-ff}.txt", DateTime.Now);
            _logger = new StreamWriter(logFolder + "\\" + logFileName, false);
            _logger.WriteLine("Log File Created: " + DateTime.Now);
        }

        /// <summary>
        /// Iterates through all directories in the given path, filtering out specific folders, 
        /// and processes each file within them based on certain conditions and file types.
        /// </summary>
        /// <param name="directoryPath"></param>
        public static void ProcessAllFoldersInDirectory(string directoryPath)
        {
            foreach (string folderPath in Directory.EnumerateDirectories(directoryPath))
            {
                if (folderPath.Contains("RPS")) { continue; }
                _folderYear = ExtractYearFromFolderName(folderPath);

                if (Convert.ToInt16(_folderYear) < 2024) { continue; }

                string folderName = Path.GetFileName(folderPath);

                Console.WriteLine("Proccessing Folder... " + folderName);
                _logger.WriteLine("Proccessing Folder... " + folderName);

                 string dmffolder = DertemineDMFFolder(folderName);
                _outputFile = Path.Combine(_outputFolder, (_folderYear + dmffolder).Trim() + ".csv");

                foreach (string folder in Directory.EnumerateDirectories(folderPath))
                {
                    if (!folder.Contains("Source")) { continue; }
                    foreach (var file in Directory.GetFiles(folder))
                    {
                        string fileName = Path.GetFileName(file);
                        WriteToLog("  Proccessing File... ", folderName, fileName);

                        if (fileName.ToUpper().Contains("EMPTY"))
                        {
                            WriteToLog("  Empty File: ", folderName, fileName);
                            continue;
                        }

                        if (fileName.ToUpper().Contains("SURVIVOR"))
                        {
                            _outputFile = Path.Combine(_outputFolder, (_folderYear + " Survivor").Trim() + ".csv");
                        }
                        else
                        {
                            dmffolder = DertemineDMFFolder(folderName);
                            _outputFile = Path.Combine(_outputFolder, (_folderYear + dmffolder).Trim() + ".csv");
                        }

                        string fileExtension = Path.GetExtension(file);
                        if (fileExtension.ToUpper() == ".CSV")
                        {
                            ProcessCSVFile(file, fileName);
                        }
                        else
                        {
                            ProcessExcelFile(file, fileName);
                        }
                    }
                }

                foreach (var file in Directory.GetFiles(folderPath))
                {
                    string fileName = Path.GetFileName(file);

                    WriteToLog("  Proccessing File... ", folderName, fileName);

                    if (fileName.ToUpper().Contains("SURVIVOR"))
                    {
                        _outputFile = Path.Combine(_outputFolder, (_folderYear + " Survivor").Trim() + ".csv");
                    }
                    else
                    {
                        dmffolder = DertemineDMFFolder(folderName);
                        _outputFile = Path.Combine(_outputFolder, (_folderYear + dmffolder).Trim() + ".csv");
                    }
                    string fileExtension = Path.GetExtension(file);
                    if (fileExtension.ToUpper() == ".CSV")
                    {
                        ProcessCSVFile(file, fileName);
                    }
                    else
                    {
                        if (_folderYear != "2011") // Skip 2011 Excell files, created CSV files for them to process
                        { ProcessExcelFile(file, fileName); }
                    }
                }
            }
        }

        /// <summary>
        /// Logs a message indicating the current file being processed.
        /// </summary>
        /// <param name="messageString"></param>
        /// <param name="folderName"></param>
        /// <param name="fileName"></param>
        private static void WriteToLog(string messageString,string folderName, string fileName)
        {
            Console.WriteLine(messageString + folderName + "\\" + fileName);
            _logger.WriteLine(messageString + folderName + "\\" + fileName);
        }

        /// <summary>
        /// Attempts to process a CSV file, handling potential IOExceptions due to file access conflicts with a retry mechanism.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="fileName"></param>
        private static void ProcessCSVFile(string file, string fileName)
        {
            bool fileProcessed = false;
            int attempt = 0;
            const int maxAttempts = 5; // Maximum number of attempts to try accessing the file
            const int delayBetweenAttempts = 3000; // Delay in milliseconds between attempts

            while (!fileProcessed && attempt < maxAttempts)
            {
                try
                {
                    using (StreamReader sr = new StreamReader(file))
                    using (StreamWriter sw = new StreamWriter(_outputFile, true))
                    {
                        int rowCount = 0;
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("Account"))
                            {
                                Console.WriteLine("  " + fileName + " has a header row.");
                                _logger.WriteLine("  " + fileName + " has a header row.");
                                continue;
                            }
                            sw.WriteLine(line);
                            ++rowCount;
                        }
                        Console.WriteLine($"  Processed {rowCount} rows");
                        _logger.WriteLine($"  Processed {rowCount} rows");
                        fileProcessed = true; // File was successfully processed
                    }
                }
                catch (IOException ex)
                {
                    attempt++;
                    Console.WriteLine($"Attempt {attempt} failed to process {fileName}: {ex.Message}");
                    _logger.WriteLine($"Attempt {attempt} failed to process {fileName}: {ex.Message}");
                    if (attempt < maxAttempts)
                    {
                        System.Threading.Thread.Sleep(delayBetweenAttempts); // Wait before retrying
                    }
                }
            }

            if (!fileProcessed)
            {
                Console.WriteLine($"  Failed to process {fileName} after {maxAttempts} attempts.");
                _logger.WriteLine($"  Failed to process {fileName} after {maxAttempts} attempts.");
            }
        }

        /// <summary>
        /// Processes an Excel file, reading the data and writing it into a CSV format.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="fileName"></param>
        private static void ProcessExcelFile(string file, string fileName)
        {
            bool fileProcessed = false;
            int attempt = 0;
            const int maxAttempts = 5; // Maximum number of attempts to try accessing the file
            const int delayBetweenAttempts = 2000; // Delay in milliseconds between attempts
            while (!fileProcessed && attempt < maxAttempts)
            {
                try
                {
                    using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                            {
                                Delimiter = ","
                            };

                            using (var csvWriter = new StreamWriter(_outputFile, true))
                            using (var csv = new CsvWriter(csvWriter, config))
                            {
                                int sheetCount = 0;
                                do
                                {
                                    Console.WriteLine($"  Processing sheet {++sheetCount}...");
                                    _logger.WriteLine($"  Processing sheet {++sheetCount}...");
                                    int rowCount = 0;
                                    while (reader.Read()) //Each ROW
                                    {
                                        var firstCellValue = reader.GetValue(0)?.ToString();
                                        if (firstCellValue != null && firstCellValue.Contains("Account"))
                                        {
                                            Console.WriteLine("  " + fileName + " has a header row.");
                                            _logger.WriteLine("  " + fileName + " has a header row.");
                                            continue;
                                        }

                                        for (int column = 0; column < reader.FieldCount; column++)
                                        {
                                            csv.WriteField(reader.GetValue(column)); //Each COLUMN
                                        }
                                        csv.NextRecord();
                                        ++rowCount;
                                    }
                                    Console.WriteLine($"  Processed {rowCount} rows");
                                    _logger.WriteLine($"  Processed {rowCount} rows");
                                    fileProcessed = true; // File was successfully processed
                                }
                                while (reader.NextResult()); //Each SHEET
                            }
                        }
                    }
                }
                catch (IOException ex)
                {
                    attempt++;
                    Console.WriteLine($"Attempt {attempt} failed to process {fileName}: {ex.Message}");
                    _logger.WriteLine($"Attempt {attempt} failed to process {fileName}: {ex.Message}");
                    if (attempt < maxAttempts)
                    {
                        System.Threading.Thread.Sleep(delayBetweenAttempts); // Wait before retrying
                    }
                }
            }

            if (!fileProcessed)
            {
                Console.WriteLine($"  Failed to process {fileName} after {maxAttempts} attempts.");
                _logger.WriteLine($"  Failed to process {fileName} after {maxAttempts} attempts.");
            }
        }

        /// <summary>
        /// Uses a regular expression to extract the year from the folder name.
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public static string ExtractYearFromFolderName(string folderPath)
        {
            Regex regex = new Regex(@"\b\d{4}\b");
            Match match = regex.Match(folderPath);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Determines the type of DMF (Death Master File) folder based on the folder name, used for organizing the output files.
        /// </summary>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static string DertemineDMFFolder(string folderName)
        {
            if (folderName.ToUpper().Contains("PENSION")) { return " Pension"; }
            if (folderName.ToUpper().Contains("GROUP")) { return " Group"; }
            return "";
        }
    }
}
