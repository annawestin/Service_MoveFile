using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;

namespace MoveQuantumFiles
{
    internal class ExcelHandler
    /// <summary>
    /// Class containing all handling with application excel.
    /// </summary>
    {
        private string _outputfolder;
        private string _nameOfFile;
        private string _fileExtension;
        private string _fileNumber;
        private readonly log4net.ILog _log;

        public ExcelHandler(log4net.ILog log)
        {
            _log = log;
        }

        public string GetFileExtension() { return _fileExtension; }

        private void SetFileExtension(string prefix) { _fileExtension = prefix; }

        public string GetOutputFolder() { return _outputfolder; }

        private void SetOutputFolder(string folder) { _outputfolder = folder; }

        public string GetNameOfFile() { return _nameOfFile; }

        private void SetNameOfFile(string name) { _nameOfFile = name; }

        public string GetFileNumber() { return _fileNumber; }

        private void SetFileNumber(string numbers) { _fileNumber = numbers; }

        public void ReadExcel(string excelPath, string sheetName, string columnName, string fileName)
        /*
         Read input excel containing where file should be moved.
        */
        {
            SetFileExtension(Path.GetExtension(fileName));

            // If filename contains numbers in a row that means they could be date
            //   or randomized numbers which will not found in input excel
            string name = fileName;
            Match match = Regex.Match(name, "[0-9]{4}?([-_0-9])*");
            if (match.Success)
            {
                _log.Info("Extracting number from name");

                // Replace numbers with wildcard
                name = fileName.Replace(match.Groups[0].Value, "%");

                // Save extracted filenumbers in object
                SetFileNumber(match.Groups[0].Value);
            }

            SetNameOfFile(Path.GetFileNameWithoutExtension(name).Trim());

            // Set filter value for SQL query
            string filterValue = GetNameOfFile() + GetFileExtension();
            if (filterValue == null)
            {
                throw new Exception("Failed to get filename of input file");
            }

            _log.Info($"Instance file number: {GetFileNumber()}, name: {GetNameOfFile()}, extension: {GetFileExtension()}");

            // Create query for filter
            string filterquery = $"SELECT * FROM[{sheetName}$] WHERE ({columnName}) LIKE ('{filterValue}')";
            _log.Info("Extracting input excel using query " + filterquery);

            // Make copy of input excel
            string inputCopy = Path.GetDirectoryName(excelPath) + "/TempCopy_" + Path.GetFileName(excelPath);
            File.Copy(excelPath, inputCopy, true);

            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={inputCopy}; Extended Properties='Excel 12.0; HDR=YES;';";

            // Connect to excel through Microsoft database engine
            OleDbConnection connection = new OleDbConnection(connectionString);
            if (connection.State == ConnectionState.Closed)
            {
                _log.Info("Connecting through MS database engine");
                connection.Open();
            }

            // Extract data from excel
            OleDbDataAdapter adp = new OleDbDataAdapter(filterquery, connection);

            // Create data table and populate with data from file
            DataTable dt = new DataTable();
            adp.Fill(dt);
            _log.Info($"{dt.Rows.Count} rows extracted from excel");

            connection.Close();
            File.Delete(inputCopy);

            if (dt.Rows.Count <= 0)
            {
                throw new Exception($"{MonitorFolder.BRE_Prefix} Could not find a match for file {fileName} in input excel file");
            }

            int rowIndex = 0;
            if (dt.Rows.Count > 1)
            {
                rowIndex = -1;
                for (int idx = 0; idx < dt.Rows.Count; idx++)
                {
                    if (dt.Rows[idx]["Filename"].ToString().Trim().Length == fileName.Length)
                    {
                        if (rowIndex < 0)
                        {
                            rowIndex = idx;
                        }
                        else
                        {
                            throw new Exception($"{MonitorFolder.BRE_Prefix} {dt.Rows.Count} possible matches found in file. Give the file a more unique name.");
                        }
                    }
                }
                if (rowIndex < 0)
                {
                    throw new Exception($"{MonitorFolder.BRE_Prefix} {dt.Rows.Count} possible matches found in file. Give the file a more unique name.");
                }
            }

            // Set new name and extension
            SetFieldsBasedDataRow(dt.Rows[rowIndex]);
        }

        private void SetFieldsBasedDataRow(DataRow row)
        /*
         Extracts the column values and set output, new file name and file extension.
        */
        {
            // Set outputfolder based on data table row
            SetOutputFolder(row["Output folder"].ToString());

            if (String.IsNullOrEmpty(GetOutputFolder()))
            {
                throw new Exception($"{MonitorFolder.BRE_Prefix} Output folder is missing in input excel file");
            }
            else
            {
                _log.Info($"Collected outputfolder: {GetOutputFolder()}");
            }

            // If file should change name after copy
            string filename;
            string columnValue = row["Name of file after copy"].ToString().Trim();
            if (columnValue.Length > 0)
            {
                // If name of file after copy should contain a date or number
                Match match = Regex.Match(columnValue, @"[Xx]{3,}|[YyMmDd-]{4,}|(\(\*\))");
                _log.Info("Setting new filename, found value to change: " + match.Value);

                if (match.Success)
                {
                    // If content with numbers/date already exist
                    if (_fileNumber != null)
                    {
                        filename = columnValue.Replace(match.Value, _fileNumber);
                        _log.Info("Set filename, with original number: " + filename);
                    }
                    // If it does not exist but it should be a date, set today's date
                    else if (match.Value.Contains("yyyy"))
                    {
                        filename = columnValue.Replace("yyyy", DateTime.Now.Year.ToString()).Replace("mm", DateTime.Now.ToString("MM")).Replace("dd", DateTime.Now.ToString("dd"));
                        _log.Info("Set filename, today's date for yyyy: " + filename);
                    }
                    else if (match.Value.Contains("YYYY"))
                    {
                        filename = columnValue.Replace("YYYY", DateTime.Now.Year.ToString()).Replace("MM", DateTime.Now.ToString("MM")).Replace("DD", DateTime.Now.ToString("dd"));
                        _log.Info("Set filename, today's date for YYYY: " + filename);
                    }
                    else
                    {
                        throw new Exception($"{MonitorFolder.BRE_Prefix} No match, adding number or date to file");
                    }
                }
                else
                {
                    filename = columnValue;
                    _log.Info("Set filename, keeping column value: " + filename);
                }
            }
            else
            {
                filename = GetNameOfFile() + GetFileNumber();
                _log.Info("Set filename, same as before copy: " + filename);
            }

            // Delete old extension
            string extension = Path.GetExtension(filename);
            if (!String.IsNullOrEmpty(extension))
            {
                filename = filename.Replace(Path.GetExtension(filename), "");
            }

            // Set file extension
            string filetypeTo = row["Change filetype to"].ToString().Trim();
            if (String.IsNullOrEmpty(filetypeTo))
            {
                _log.Info("Using same extension as input file: " + GetFileExtension());
            }
            else
            {
                if (!filetypeTo.Contains("."))
                {
                    filetypeTo = "." + filetypeTo;
                }
                SetFileExtension(filetypeTo);
                _log.Info("Set prefix, from column value: " + GetFileExtension());
            }

            // Set new name of file with new prefix
            SetNameOfFile(filename + GetFileExtension());
            _log.Info("Set filename: " + GetNameOfFile());
        }


    }
}