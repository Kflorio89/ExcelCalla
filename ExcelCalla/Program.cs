using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ExcelCalla
{
    class Program
    {
        private static readonly ILog _log = LogManager.GetLogger("ExcelCalla.log");
        // Order: Tube Height(6), Width(8), Area Total(10), Area Top(12), Flatness(14), Max Curvature(16), 
        // Slope Ave(18), Slope X Ave(20), Slope R Ave(22), Slope Width(24), Recirculation Ave(26) 
        private static readonly int StartPosition = 6;
        private static readonly int DataCount = 11;
        private static readonly string SheetName = "SingleBall";
        private static int StartLineNumber = 1;
        private static string txtFileFolderPath;// = @"C:\Users\kflor\OneDrive\Desktop\Streamline\Parts\Golf Ball\Jobs\11212220_123055_D8807_001_Stage01";
        private static string excelFilePath;// = @"C:\Users\kflor\OneDrive\Desktop\Streamline\Parts\Golf Ball\Archive\SingleBall Report Template.xlsx";

        /// <summary>
        /// Main entry point of program
        /// </summary>
        /// <param name="args"> Expecting 2 arguments, Argument #1: Path to folder with all the .stats files for a single ball,
        /// Argument #2: excel template file location to make a copy of and place in desired location to fill out. </param>
        static void Main(string[] args)
        {
            _log.Info("Starting Executable..");
            // Command line argument section
            #region Command line

            _log.Info($"Number of arguments passed for this execution: {args.Length}");
            for (int i = 0; i < args.Length; ++i)
            {
                _log.Info($"Arg #{i + 1}: {args[i]}");
            }

            if (args.Length < 1)
            {
                _log.Error("No arguments passed.");
                return;
            }

            if (args.Length < 2)
            {
                _log.Error("Only one argument passed.");
                return;
            }

            // Should be at least 2 arguments which will be the paths to the two files
            excelFilePath = args[0];
            txtFileFolderPath = args[1];

            if (!File.Exists(excelFilePath))
            {
                _log.Error($"File does not exist or path invalid: {excelFilePath}");
                return;
            }

            if (!Directory.Exists(txtFileFolderPath))
            {
                _log.Error($"Directory does not exist or invalid: {txtFileFolderPath}");
                return;
            }
            #endregion Command line

            // Split folder name to get: 0 = Date, 1 = Time, 2 = Group#, 3 = ball#, 4 = stage#
            string[] info = Path.GetFileName(txtFileFolderPath).Split('_');

            if (info.Length < 4)
            {
                _log.Error($"Error with number of details extracted from filename {txtFileFolderPath}.");
            }
            string archivefolder = Path.GetDirectoryName(excelFilePath);
            string golfBallFolder = Path.GetDirectoryName(archivefolder);
            string ballFolder = golfBallFolder + $"\\Reports\\{info[0]}_{info[1]}_{info[4]}";

            // Check if folder exists with the correct name, if not then create it
            if (!Directory.Exists(ballFolder))
            {
                Directory.CreateDirectory(ballFolder);
            }

            string newFileName = $"{ballFolder}\\{info[2]}_{info[3]}_{info[4]}.xlsx";
            if (File.Exists(newFileName))
            {
                _log.Error($"File already exists {newFileName}.");
            }
            try
            {
                // Copy Template and rename and place in Directory - ballFolder
                File.Copy(excelFilePath, newFileName);
            }
            catch (Exception ex)
            {
                _log.Error($"Exception when copying file from {excelFilePath} as {newFileName}, error message: {ex.Message.ToString()}");
            }
            // Open initial Directory and load all file names that have .stats
            List<string> dataFileNames = new List<string>();
            try
            {
                DirectoryInfo direct = new DirectoryInfo(txtFileFolderPath);
                FileInfo[] files = direct.GetFiles("*.stats");

                foreach (FileInfo inf in files)
                {
                    dataFileNames.Add(inf.FullName);
                }

                if (dataFileNames.Count == 0)
                {
                    _log.Error($"No Data files inside {txtFileFolderPath}");
                    return;
                }
            }
            catch (Exception ex)
            {
                _log.Error($"Issue getting file names from directory {txtFileFolderPath}, message : {ex.Message.ToString()}");
                return;
            }

           //create a fileinfo object of an excel file on the disk (file must exist)
            FileInfo file;
            try
            {
                file = new FileInfo(excelFilePath);
            }
            catch (Exception ex)
            {
                _log.Error($"Exception thrown when creating file object of excel file. {ex.Message.ToString()}");
                return;
            }

            #region Create Excel library object
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //create a new Excel package from the file
            using (ExcelPackage excelDoc = new ExcelPackage(file))
            {
                if (excelDoc.Workbook.Worksheets.FirstOrDefault(x => x.Name == SheetName) == null)
                {
                    _log.Error("Worksheet does not exist..");
                    return;
                }

                //Get a WorkSheet by name. Note that EPPlus indexes start at 1!
                ExcelWorksheet firstWorksheet = excelDoc.Workbook.Worksheets[SheetName];

                // Find first unused row with no data entered
                while (firstWorksheet.Cells[StartLineNumber, 1].Value == null || !firstWorksheet.Cells[StartLineNumber, 1].Value.ToString().ToLower().Contains("geometry"))
                {
                    ++StartLineNumber;
                }
                _log.Info($"Geometry #s start at worksheet row: {++StartLineNumber}");

                #region Text File Data Retrieval
                for (int fileIndex = 0; fileIndex < dataFileNames.Count; ++fileIndex)
                {
                    // Read in text file section
                    int lineNumber = StartLineNumber;
                    string GeoNumber = "";
                    string path = dataFileNames[fileIndex];
                    string fileName = Path.GetFileName(path).Split('.')[0];
                    double[] dataEntries;

                    try
                    {
                        dataEntries = ParseTextFile(path, out GeoNumber);
                    }
                    catch (Exception ex)
                    {
                        _log.Error($"Issue parsing textfile in path: {path}, error message: {ex.Message.ToString()}");
                        return;
                    }

                    // Find line with corresponding geometry number
                    while ((string)firstWorksheet.Cells[lineNumber, 1].Value != GeoNumber)
                    {
                        ++lineNumber;
                    }

                    // Insert data into sheet
                    for (int i = 0; i < DataCount; ++i)
                    {
                        firstWorksheet.Cells[lineNumber, StartPosition + i].Value = dataEntries[i];
                    }

                }
                #endregion

                // Save the changes
                try
                {
                    excelDoc.Save();
                }
                catch (Exception ex)
                {
                    _log.Error($"Exception when saving excel document: {ex.Message.ToString()}");
                }
            }
            #endregion
        }

        /// <summary>
        /// Parse text file passed and pulls data out
        /// </summary>
        /// <param name="path"> Path to the text file </param>
        /// <param name="geoNum"> The geometry number pulled from text file name </param>
        /// <returns> An array of double values from data sheet </returns>
        private static double[] ParseTextFile(string path, out string geoNum)
        {
            // Pull Geometry number from text file name
            geoNum = Path.GetFileName(path).Split('_')[3].Split('.')[0];

            // Open text file and read all lines
            double[] data = new double[DataCount];

            // Read all lines from datafile and pull values out, format should be: DataName=Value
            string[] temp = File.ReadAllLines(path);
            for (int i = 0; i < DataCount; ++i)
            {
                if (Double.TryParse(temp[i].Split('=')[1], out double num))
                {
                    data[i] = num;
                }
                else
                {
                    _log.Error($"Issue when pulling specific data entry from text file on line number {i + 1}");
                }
            }

            return data;
        }
    }
}
