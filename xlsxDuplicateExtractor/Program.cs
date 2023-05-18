using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace xlsxDuplicateExtractor
{
    class Program
    {
        const int columnsToSkip = 3;
        const int minAmmountOfRepeats = 3;

        static void Main(string[] args)
        {
            var inputFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "Input");
            var outputFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "Output");

            if (!Directory.Exists(inputFolderPath))
            {
                Console.WriteLine($"Input folder '{inputFolderPath}' does not exist.");
                return;
            }

            if (!Directory.Exists(outputFolderPath))
            {
                Directory.CreateDirectory(outputFolderPath);
            }

            var inputFiles = Directory.GetFiles(inputFolderPath, "*.xlsx");

            if (inputFiles.Length == 0)
            {
                Console.WriteLine("No input files found in the 'Input' folder.");
                return;
            }

            foreach (var inputFilePath in inputFiles)
            {
                var inputFileName = Path.GetFileName(inputFilePath);
                var outputFilePath = Path.Combine(outputFolderPath, inputFileName);

                Console.WriteLine($"Processing '{inputFileName}'...");

                try
                {

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage inputPackage = new ExcelPackage(new FileInfo(inputFilePath)))
                    using (ExcelPackage outputPackage = new ExcelPackage(new FileInfo(outputFilePath)))
                    {
                        ExcelWorksheet worksheet = inputPackage.Workbook.Worksheets[0];

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // Key: Concatenated string from 4th column to the end. Value: Row number.
                        var rowMap = new Dictionary<string, List<int>>();
                        // Key: Row number. Value: Row content.
                        var repeatValueRows = new Dictionary<int, string>();

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var rowString = string.Empty;
                            for (int col = columnsToSkip; col <= colCount; col++)
                            {
                                var cellValue = Convert.ToString(worksheet.Cells[row, col].Value);
                                rowString += cellValue;
                            }

                            // Add the row number to the list in the rowMap
                            if (rowMap.ContainsKey(rowString))
                                rowMap[rowString].Add(row);
                            else
                                rowMap[rowString] = new List<int> { row };
                        }

                        // List of rows with any value present 3 or more times.
                        var repeatedValueRows = new List<int>();

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var rowString = string.Empty;
                            var cellValues = new Dictionary<string, int>();
                            for (int col = columnsToSkip; col <= colCount; col++)
                            {
                                var cellValue = Convert.ToString(worksheet.Cells[row, col].Value);
                                rowString += cellValue;

                                // Ignore empty strings or null values
                                if (string.IsNullOrWhiteSpace(cellValue))
                                    continue;

                                // Count the occurrence of cell value in the current row
                                if (cellValues.ContainsKey(cellValue))
                                    cellValues[cellValue]++;
                                else
                                    cellValues[cellValue] = 1;
                            }

                            // Check if any non-empty cell value occurred three or more times in the current row
                            // and it is not a duplicate row
                            if (cellValues.Any(pair => pair.Value >= minAmmountOfRepeats) && rowMap[rowString].Count == 1)
                            {
                                repeatedValueRows.Add(row);
                            }
                        }

                        // Prepare data for writing to new excel file
                        var duplicateRows = rowMap.Where(pair => pair.Value.Count > 1)
                                                  .SelectMany(pair => pair.Value)
                                                  .ToList();

                        // Write to new excel file
                        using (outputPackage)
                        {
                            var worksheet1 = outputPackage.Workbook.Worksheets.Add("Duplicates");
                            var worksheet2 = outputPackage.Workbook.Worksheets.Add("Repeats");

                            WriteRowsToWorksheet(duplicateRows, worksheet, worksheet1, rowMap);
                            WriteRowsToWorksheet(repeatedValueRows, worksheet, worksheet2);

                            outputPackage.Save();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while processing '{inputFileName}': {ex.Message}");
                }
            }
        }
        private static void WriteRowsToWorksheet(List<int> rows, ExcelWorksheet sourceWorksheet, ExcelWorksheet targetWorksheet, Dictionary<string, List<int>> rowMap = null)
        {
            int colCount = sourceWorksheet.Dimension.Columns;

            // Write header
            for (int col = 1; col <= colCount; col++)
            {
                targetWorksheet.Cells[1, col].Value = sourceWorksheet.Cells[1, col].Value;
            }

            var sortedRows = rows;

            // Sort rows
            // If the rowMap is provided, sort by duplicate count and then by 3rd column
            if (rowMap != null)
            {
                sortedRows = rows.GroupBy(row => rowMap.Values.First(v => v.Contains(row)))
                                 .OrderByDescending(g => g.Key.Count)
                                 .ThenBy(g => sourceWorksheet.Cells[g.First(), 3].Value)
                                 .SelectMany(g => g)
                                 .ToList();
            }
            else
            {
                sortedRows = rows.OrderBy(row => sourceWorksheet.Cells[row, columnsToSkip].Value.ToString()).ToList();
            }

            // Write rows
            for (int i = 0; i < sortedRows.Count; i++)
            {
                int sourceRow = sortedRows[i];
                int targetRow = i + 2; // +1 for zero-indexing, +1 for header
                for (int col = 1; col <= colCount; col++)
                {
                    targetWorksheet.Cells[targetRow, col].Value = sourceWorksheet.Cells[sourceRow, col].Value;
                }
            }
        }
    }
}