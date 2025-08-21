using System;
using System.Configuration;
using System.IO;
using ClosedXML.Excel;
using EmsConsole;

namespace CsvToExcelOrganizer
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string vEmsCsvFiles = ConfigurationManager.AppSettings["CsvFiles"], vEmsDirectory = ConfigurationManager.AppSettings["ExcelFiles"], connectionString = ConfigurationManager.AppSettings["SqlConnectionString"];
                SqlData sqlData = new SqlData(connectionString);

                var today = DateTime.Today;

                if (!Directory.Exists(vEmsCsvFiles))
                {
                    Console.WriteLine($"Source folder not found: {vEmsCsvFiles}");
                    return;
                }

                // Create Year/Month/Date folder structure
                string path = Path.Combine(vEmsDirectory, today.Year.ToString());
                Directory.CreateDirectory(path);

                path = Path.Combine(path, today.ToString("MM"));
                Directory.CreateDirectory(path);

                string dateFolderPath = Path.Combine(path, today.ToString("ddMMyyyy"));
                Directory.CreateDirectory(dateFolderPath);

                // Copy today's CSV files
                foreach (var oFilePath in Directory.EnumerateFiles(vEmsCsvFiles, "*.csv"))
                {
                    var fileDate = File.GetLastWriteTime(oFilePath).Date;
                    if (fileDate == today)
                    {
                        string vDirectoryPath = Path.Combine(dateFolderPath, Path.GetFileName(oFilePath));
                        if (!File.Exists(vDirectoryPath))
                            File.Copy(oFilePath, vDirectoryPath);
                    }
                }

                // Convert CSV → Excel
                foreach (var oCsvFile in Directory.EnumerateFiles(dateFolderPath, "*.csv"))
                {
                    string excelFilePath = Path.ChangeExtension(oCsvFile, ".xlsx");
                    if (!File.Exists(excelFilePath))
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int row = 1;
                            using (var reader = new StreamReader(oCsvFile))
                            {
                                while (!reader.EndOfStream)
                                {
                                    var values = reader.ReadLine().Split(',');
                                    for (int col = 0; col < values.Length; col++)
                                        worksheet.Cell(row, col + 1).Value = values[col];
                                    row++;
                                }
                            }
                            workbook.SaveAs(excelFilePath);
                        }
                    }

                    // Insert Excel data into SQL
                    try
                    {
                        sqlData.ProcessFile(excelFilePath);
                    }
                    catch (Exception o)
                    {
                        Console.WriteLine($"Error inserting SQL for {excelFilePath}: {o.Message}");
                    }
                }

                Console.WriteLine("All processing completed successfully.");
            }
            catch (Exception o)
            {
                Console.WriteLine($"Unexpected error: {o.Message}");
            }
        }
    }
}
