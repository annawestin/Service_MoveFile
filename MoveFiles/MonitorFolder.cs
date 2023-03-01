using System;
using System.IO;
using System.Collections;
using System.Threading.Tasks;

namespace MoveQuantumFiles
{
    internal class MonitorFolder
    /// <summary>
    /// Monitors folder specified in config.
    /// </summary>
    {
        private const string TimeStampFormat = "yyyy-MM-dd HH:mm";
        public const string BRE_Prefix = "ERR:"; // these exceptions will skip the technical parts and only log our defined descriptions
        private readonly string _directoryToWatch;
        private readonly string _errorFolder;
        private readonly string _heartBeatFile;
        private readonly log4net.ILog _log;
        private readonly Hashtable _config;
        private string _lastTimeStr = "<no files processed yet>";
        private string _currentError = "";

        public MonitorFolder(log4net.ILog log, Hashtable configSettings)
        {
            _log = log;
            _config = configSettings;
            _directoryToWatch = configSettings["MonitorPath"].ToString().Trim();
            if (_directoryToWatch.EndsWith("\\") || _directoryToWatch.EndsWith("/"))
            {
                _directoryToWatch = _directoryToWatch.Substring(0, _directoryToWatch.Length - 1);
            }
            _errorFolder = _directoryToWatch + "\\Error";
            _heartBeatFile = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Heartbeat_MoveFiles.txt";
        }

        public void WriteConfigSettings()
        {
            _log.Info($"Excel path: " + _config["InputExcel"].ToString());
            _log.Info($"Excel sheet name: " + _config["InputExcelSheetName"].ToString());
            _log.Info($"Excel column name: " + _config["InputExcelColumnName"].ToString());
            _log.Info($"Watching folder: " + _directoryToWatch);
        }

        public void ValidateFolders()
        {
            if (!Directory.Exists(_directoryToWatch))
            {
                _currentError = "Can't find main folder: " + _directoryToWatch;
                throw new Exception(_currentError);
            }

            if (!Directory.Exists(_errorFolder))
            {
                Directory.CreateDirectory(_errorFolder);
                System.Threading.Thread.Sleep(3000); // wait to be sure the network call is done

                if (!Directory.Exists(_errorFolder))
                {
                    _currentError = "Can't find or create error folder: " + _errorFolder;
                    throw new Exception(_currentError);
                }
            }
            _currentError = "";
        }

        public void CheckForNewFiles()
        {
            try
            {
                var dirInfo = new DirectoryInfo(_directoryToWatch);
                var files = dirInfo.GetFiles("*.*", SearchOption.TopDirectoryOnly);
                bool processedFiles = false;

                System.Threading.Thread.Sleep(5000); // wait for any file changes to be ready

                foreach (FileInfo file in files)
                {
                    if (ProcessFile(file.FullName, _errorFolder))
                    {
                        processedFiles = true;

                    }
                }

                if (processedFiles)
                {
                    _lastTimeStr = DateTime.Now.ToString(TimeStampFormat);
                    ExcelHandler statexcel = new ExcelHandler(_log);
                }

                WriteHeartbeatFile();
            }
            catch (DirectoryNotFoundException)
            {
                _log.Error($"Can't find the folder: " + _directoryToWatch);
            }
            catch (Exception ex)
            {
                _log.Info("-- Catch CheckForNewFiles --");
                _log.Error(ex);
            }
        }

        private bool ProcessFile(string fullFileName, string errorFolder)
        {
            if (fullFileName.Contains("ExceptionReport") || fullFileName.Contains("PTR001A") || !File.Exists(fullFileName))
            {
                return false;
            }

            _log.Info($"File found to process: {fullFileName}");

            ExcelHandler excel = new ExcelHandler(_log);

            try
            {
                string excelpath = _config["InputExcel"].ToString();
                string sheetName = _config["InputExcelSheetName"].ToString();
                string columnName = _config["InputExcelColumnName"].ToString();
                string filename = Path.GetFileName(fullFileName);

                // The database connection can freeze the process for days,
                //   and therefore it's runned as a task because a Task object
                //   is immediately returned by the Task.Run method
                var task = Task.Run(() =>
                {
                    excel.ReadExcel(excelpath, sheetName, columnName, filename);
                });

                // If the task is successfully completed within 45 seconds
                if (!task.Wait(TimeSpan.FromMilliseconds(45000)))
                {
                    throw new TimeoutException("Reading of Data sheet file has taken longer than the maximum time allowed.");
                }

                // Copy file
                string outputFolder = excel.GetOutputFolder();
                if (outputFolder == null)
                {
                    throw new Exception($"{BRE_Prefix} No output folder found for file {filename}");
                }
                else if (!Directory.Exists(outputFolder))
                {
                    throw new Exception($"{BRE_Prefix} Directory {outputFolder} does not exist");
                }

                string newfileName = excel.GetNameOfFile();
                if (newfileName == null)
                {
                    throw new Exception($"Something went wrong when setting new filename for file {filename}");
                }

                MoveFile(fullFileName, outputFolder, newfileName);
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith(BRE_Prefix))
                {
                    _log.Error(ex.Message);
                }
                else if (ex.InnerException != null && ex.InnerException.Message.StartsWith(BRE_Prefix))
                {
                    _log.Error(ex.InnerException.Message);
                }
                else
                {
                    _log.Info("-- Catch ProcessFile --");
                    _log.Error(ex);
                }

                try
                {
                    string uniqueFileName = Path.GetFileName(fullFileName);
                    if (File.Exists(errorFolder + "\\" + uniqueFileName))
                    {
                        uniqueFileName = Path.GetFileNameWithoutExtension(uniqueFileName) + "_" + DateTime.Now.ToString("yyMMdd_HHmmss") + Path.GetExtension(uniqueFileName);
                    }
                    MoveFile(fullFileName, errorFolder, uniqueFileName);
                }
                catch (Exception moveEx)
                {
                    _log.Info("-- Catch ProcessFile, move file to error folder --");
                    _log.Error(moveEx.Message);
                }
            }
            return true; // the file has been processed, with or without an error
        }

        private void MoveFile(string inputfile, string outputfolder, string filename)
        {
            filename = outputfolder + "\\" + filename;
            _log.Info($"Moving file {inputfile} to {filename}");

            File.Copy(inputfile, filename, true);
            System.Threading.Thread.Sleep(3000);
            if (!File.Exists(filename))
            {
                throw new Exception($"{BRE_Prefix} Couldn't move {inputfile} to {filename}");
            }

            File.Delete(inputfile);
            int count = 3;
            do
            {
                if (!File.Exists(inputfile))
                {
                    return;
                }
                // Wait and check again if the the network is slow
                System.Threading.Thread.Sleep(3000);
                count--;
            }
            while (count > 0);

            throw new Exception($"{BRE_Prefix} Couldn't delete {inputfile}");
        }


        public void WriteHeartbeatFile()
        {
            if (File.Exists(_heartBeatFile))
            {
                File.Delete(_heartBeatFile);
            }

            using (StreamWriter sw = File.CreateText(_heartBeatFile))
            {
                sw.WriteLine($"Heartbeat: {DateTime.Now.ToString(TimeStampFormat)}");
                sw.WriteLine($"Last time handling files: {_lastTimeStr}");
                sw.WriteLine($"Directory watched: {_directoryToWatch}");
                if (!String.IsNullOrEmpty(_currentError))
                {
                    sw.WriteLine($"Current error: {_currentError}");
                }
            }
        }
    }
}
