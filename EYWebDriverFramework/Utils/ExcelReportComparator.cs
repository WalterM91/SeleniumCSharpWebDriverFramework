using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EYWebDriverFramework.Config;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace EYWebDriverFramework.Utils
{
    public class ExcelReportComparator
    {
        public static Dictionary<string, Cell[,]> expectedExcel;
        public static Dictionary<string, Cell[,]> actualExcel;

        readonly string goldenCopyPath;
        List<CellComment> commentList;
        bool result;
        int error;

        public ExcelReportComparator(string goldenCopyPath)
        {
            this.goldenCopyPath = goldenCopyPath;
        }

        public bool CompareExcelReport(string actualPath)
        {
            Excel.Application app =null;
            Excel.Workbook goldenCopy=null;
            Excel.Workbook actualWorkbook=null;
            var errorList = ExcelComparator.CompareExcelDataTable(goldenCopyPath, actualPath);
            try
            {
                app = new Excel.Application
                {
                    DisplayAlerts = false,
                    ScreenUpdating = false
                };
                goldenCopy = app.Workbooks.Open(goldenCopyPath, Notify: false);
                commentList = new List<CellComment>();
                result = true;

                //Open files to compare
                actualWorkbook = app.Workbooks.Open(actualPath, Notify: false);
                foreach (Excel.Worksheet expectedWS in goldenCopy.Worksheets)
                {
                    // Discard Hidden tabs
                    if (!expectedWS.Visible.Equals(Excel.XlSheetVisibility.xlSheetVisible))
                        continue;
                    Excel.Worksheet actualWS = null;
                    try
                    {
                        actualWS = actualWorkbook.Worksheets[expectedWS.Name];
                        var worksheetErrorCells = errorList[expectedWS.Name];
                        LogHelpers.Write("Starting Review of potential errors sheet: " + expectedWS.Name);
                        // Find the last real row and column that has a value.
                        Stopwatch stopWatch = new Stopwatch();
                        stopWatch.Start();
                        //Checks that cells matches
                        int total = worksheetErrorCells.Count;
                        int perc = 0;
                        error = 0;
                        foreach (var cell in worksheetErrorCells)
                        {
                            int i = cell.Item1;
                            int j = cell.Item2;
                            ExcelCell expected = new ExcelCell(expectedWS.Cells[i, j]);
                            ExcelCell actual = new ExcelCell(actualWS.Cells[i, j]);
                            if(!CellsAreEqual(expected, actual))
                            {
                                actualWS.MarkCellAsError(i, j);
                                ComparatorError(i, j, expected.GetValueWithFormat(), actual.GetValueWithFormat());
                            }
                            perc++;
                        }
                        // Add error cell comments
                        foreach (CellComment c in commentList)
                        {
                            if (c.text != null)
                                actualWS.Cells[c.row, c.col].AddComment(c.text);
                        }
                        commentList.Clear();
                        if (total != 0)
                            LogHelpers.Write("Completed review: " + (perc * 100 / total) + "% " + perc + "/" + total + " Errors found:" + error + " Elapsed Time: " + stopWatch.Elapsed.TotalSeconds);
                        else
                            LogHelpers.Write("Completed review: no potential errors found");
                    }
                    catch (COMException)
                    {
                        // The whole sheet is missing,  i add it to the actual result and marked as diff.
                        LogHelpers.Write("Missing Worksheet " + expectedWS.Name);
                        Excel.Worksheet missingWS = actualWorkbook.Worksheets.Add();
                        missingWS.Name = expectedWS.Name;
                        missingWS.Tab.Color = Excel.XlRgbColor.rgbRed;
                        missingWS.MarkCellAsError(1, 1);
                        result = false;
                    }
                    finally
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(expectedWS);
                        if (actualWS!=null)
                            Marshal.ReleaseComObject(actualWS);
                    }
                }
                //Close workbooks
                if (!result)
                {
                    string fileName = string.Format("Result_{0}_{1}", DateTime.Now.Add(default).ToString("MM-dd-yyyy_HHmmss", new CultureInfo("en-US")), actualWorkbook.Name);

                    string resultPath = ConfigReader.TestResultsPath + fileName;
                    actualWorkbook.SaveAs(resultPath);
                    LogHelpers.Write(string.Format("Comparison report generated in the following path: {0}", resultPath));
                    ConfigReader.reportGenerated = resultPath;
                }
            }
            finally {
                if(actualWorkbook!=null)
                    actualWorkbook.Close(false);
                if (goldenCopy != null)
                    goldenCopy.Close(false);
                if (app != null)
                    app.Quit();
                Marshal.ReleaseComObject(app);
            }

            if (!result)
            {

                TestContext.WriteLine("3rd spreadsheet (if present) shows the differences between the downloaded report and the golden copy");
                TestContext.AddTestAttachment(ConfigReader.reportGenerated, "Differences");
            }
            if(ExcelComparator.ex != null)
            {
                throw ExcelComparator.ex;
            }
            return result;
        }

        /// <summary>
        /// Compares both files by using OpenXML SDK.
        /// </summary>
        /// <param name="actualPath">The path of the file that was downloaded from Allokate.</param>
        /// <returns>A bool indicating wether comparison was successful (no differences found) or not.</returns>
        public bool DevopsCompareExcelReport(string actualPath)
        {
            PrepareExcelFiles(actualPath, 
                out string recoveredGoldenPath, 
                out SpreadsheetDocument recoveredGoldenFile, 
                out SpreadsheetDocument recoveredActualFile);

            var errorList = ExcelComparator.CompareExcelDataTable(goldenCopyPath, actualPath);

            result = true;
            commentList = new List<CellComment>();

            foreach (Sheet expectedSheet in recoveredGoldenFile.WorkbookPart.Workbook.Sheets)
            {
                if (expectedSheet.State != null && expectedSheet.State != SheetStateValues.Visible)
                    continue;

                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                WorksheetPart expectedWorksheetPart = recoveredGoldenFile.WorkbookPart.GetPartById(expectedSheet.Id.Value) as WorksheetPart;
                Worksheet expectedWorksheet = expectedWorksheetPart.Worksheet;

                WorksheetPart actualWorksheetPart = GetActualWorksheetPart(recoveredGoldenFile, recoveredActualFile, expectedSheet);

                Worksheet actualWorksheet = actualWorksheetPart.Worksheet;

                //The structuration of attributes is made only in those worksheets that are going to be compared
                ExcelReader.StructurateExcelSheetDataXmlAttributes(recoveredActualFile, actualWorksheet);
                ExcelReader.StructurateExcelSheetDataXmlAttributes(recoveredGoldenFile, expectedWorksheet);

                var worksheetErrorCells = errorList[expectedSheet.Name];

                LogHelpers.Write("Starting Review of potential errors sheet: " + expectedSheet.Name);

                int total = worksheetErrorCells.Count;
                int perc = 0;
                error = 0;
                foreach (var cell in worksheetErrorCells)
                {
                    int rowId = cell.Item1;
                    int colId = cell.Item2;

                    ExcelCell expectedCell = new ExcelCell(ExcelReader.GetCell(expectedWorksheet, rowId, colId), recoveredGoldenFile);
                    ExcelCell actualCell = new ExcelCell(ExcelReader.GetCell(actualWorksheet, rowId, colId), recoveredActualFile);

                    if (!expectedCell.CompareWithXML(actualCell))
                    {
                        recoveredGoldenFile.MarkCellAsError(expectedWorksheet, ExcelReader.GetCell(expectedWorksheet, rowId, colId));
                        ComparatorError(rowId, colId, expectedCell.GetValueWithFormatXML(), actualCell.GetValueWithFormatXML());
                    }
                    perc++;
                }

                ExcelReader.InsertComments(expectedWorksheetPart, commentList);

                stopWatch.Stop();

                commentList.Clear();
                if (total != 0)
                    LogHelpers.Write("Completed review: " + (perc * 100 / total) + "% " + perc + "/" + total + " Errors found:" + error + " Elapsed Time: " + stopWatch.Elapsed.TotalSeconds);
                else
                    LogHelpers.Write("Completed review: no potential errors found");

            }

            recoveredActualFile.Dispose();
            recoveredGoldenFile.Dispose();

            if (!result)
            {
                TestContext.WriteLine("3rd spreadsheet (if present) shows the differences between the downloaded report and the golden copy");

                recoveredGoldenPath = PrepareDifferencesFile(recoveredGoldenPath);

                TestContext.AddTestAttachment(recoveredGoldenPath, "Differences");
                LogHelpers.Write(string.Format("Comparison report generated in the following path: {0}", recoveredGoldenPath));
            }

            if (ExcelComparator.ex != null)
            {
                throw ExcelComparator.ex;
            }

            return result;
        }

        private static string PrepareDifferencesFile(string recoveredGoldenPath)
        {
            //To make it easier for those who read test results if comparison wasn't successful the name of golden Copy
            //structurated file is changed to "fileName - Differences.xlsm" to attach it to test results.
            FileInfo fileInfo = new FileInfo(recoveredGoldenPath);
            recoveredGoldenPath = recoveredGoldenPath.Replace("Structured", "Differences");
            File.Move(fileInfo.FullName, fileInfo.FullName.Replace("Structured", "Differences"));
            return recoveredGoldenPath;
        }

        private static WorksheetPart GetActualWorksheetPart(SpreadsheetDocument recoveredGoldenFile, SpreadsheetDocument recoveredActualFile, Sheet expectedSheet)
        {
            WorksheetPart actualWorksheetPart;

            try
            {
                actualWorksheetPart = recoveredActualFile.WorkbookPart.GetPartById(expectedSheet.Id.Value) as WorksheetPart;
            }
            catch (ArgumentOutOfRangeException e)
            {
                ExcelComparator.ex = new ArgumentOutOfRangeException("Missing worksheet: " + expectedSheet.Name.Value, e);
                bool sheetFound = false;
                string sheetId = string.Empty;

                foreach (Sheet sheet in recoveredActualFile.WorkbookPart.Workbook.Sheets)
                {
                    if (Equals(sheet.Name.Value, expectedSheet.Name.Value))
                    {
                        sheetId = sheet.Id.Value;
                        sheetFound = true;
                        break;
                    }
                }
                if (!sheetFound)
                {
                    recoveredActualFile.Dispose();
                    recoveredGoldenFile.Dispose();
                    throw ExcelComparator.ex;
                }
                else
                {
                    actualWorksheetPart = recoveredActualFile.WorkbookPart.GetPartById(sheetId) as WorksheetPart;
                }
            }

            return actualWorksheetPart;
        }

        private void PrepareExcelFiles(string actualPath, out string recoveredGoldenPath, out SpreadsheetDocument recoveredGoldenFile, out SpreadsheetDocument recoveredActualFile)
        {
            SpreadsheetDocument actualFile = ExcelReader.OpenExcelFile(actualPath, false);
            SpreadsheetDocument goldenFile = ExcelReader.OpenExcelFile(goldenCopyPath, false);

            //Files are restructurated by adding all RowIndex and CellReference attributes. For that, new paths are created based on originals
            //for both Actual and Golden files at Settings.DownloadPath location to place there those restructurated files.
            //Also " - Structured" is added to file name to separate them from originals.
            FileInfo goldenFileInfo = new FileInfo(goldenCopyPath);
            FileInfo actualFileInfo = new FileInfo(actualPath);
            recoveredGoldenPath = Path.Combine(Settings.DownloadPath, goldenFileInfo.Name.Replace(".xlsm", " - Structured.xlsm"));
            string recoveredActualPath = Path.Combine(Settings.DownloadPath, actualFileInfo.Name.Replace(".xlsm", " - Structured.xlsm"));

            //Files are cloned from originals and opened in read/write mode so structuration can be saved.
            recoveredGoldenFile = (SpreadsheetDocument)goldenFile.Clone(recoveredGoldenPath, isEditable: true);
            recoveredActualFile = (SpreadsheetDocument)actualFile.Clone(recoveredActualPath, isEditable: true);
            actualFile.Dispose();
            goldenFile.Dispose();
        }

        private bool CellsAreEqual(ExcelCell expected, ExcelCell actual)
        {
            if (expected.Value == null)
            {
                if (actual.Value != null)
                {
                    return false;
                }
            }
            else if (!expected.CompareWith(actual))
            {
                return false;
            }
            return true;
        }

        private void ComparatorError(int i, int j, string expected, string actual)
        {
            LogHelpers.Write(string.Format("ERROR in cell {0}. Expected: \"{1}\" Actual: \"{2}\"", ExcelComparator.GetCellName(i,j), expected, actual));
            commentList.Add(new CellComment(i, j, string.Concat("Golden Copy Value:",
                Environment.NewLine, expected, Environment.NewLine, "Downloaded Report Value:", Environment.NewLine, actual)));
            error++;
            result = false;
        }
    }

    internal static class RangeExtensions
    {
        public static void MarkCellAsError(this SpreadsheetDocument workdocument, Worksheet worksheet, Cell expectedCell)
        {
            StyleSheetFunctions.GetAndOrSetErrorStyleID(workdocument, worksheet, expectedCell);
            workdocument.Save();
        }
        public static void MarkCellAsError(this Excel.Worksheet ws, int i, int j)
        {
            // if a cell value is missing the comparison fails and I marked the diff with red
            Excel.Range range = (ws.Cells[i, j] as Excel.Range);
            range.Interior.Color = Excel.XlRgbColor.rgbRed;
        }
    }
}
