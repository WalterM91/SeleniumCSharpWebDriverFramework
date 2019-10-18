using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EYWebDriverFramework.Config;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace EYWebDriverFramework.Utils
{
    public class ExcelCell
    {
        public dynamic Value { get; private set; }
        private string Format { get; set; }
        public ExcelCell(Excel.Range cell)
        {
            Value = cell.Value;
            Format = Convert.ToString(cell.NumberFormat);
        }
        public ExcelCell(Cell cell, SpreadsheetDocument document)
        {
            Value = ExcelReader.GetCellValue(document, cell);
            Format = GetCellFormatCode(document, cell);
        }

        public string GetValueWithFormat()
        {
            return CellFormatter.ApplyFormat(Value, Format);
        }

        public string GetValueWithFormatXML()
        {
            return CellFormatter.XmlApplyFormat(Value, Format);
        }

        public bool CompareWith(ExcelCell actual)
        {
            return CellComparator.CompareCells(Value, actual.Value, Format);
        }

        public bool CompareWithXML(ExcelCell actual)
        {
            return CellComparator.CompareCellsXML(Value, actual.Value, Format);
        }

        private string GetCellFormatCode(SpreadsheetDocument document, Cell cell)
        {
            if (CellDoesNotHaveStyleIndexAttribute(cell))
            {
                return "General";
            }
            else
            {
                GetStyleSheet(document, out Stylesheet currentDocumentStylesheet);
                GetCurrentCellFormatNode(cell, currentDocumentStylesheet, out CellFormat currentCellFormatNode);

                GetCurrentCellNumberFormatIdNode(currentDocumentStylesheet, currentCellFormatNode, out NumberingFormat currentCellNumberFormatId);

                if (IsFormatPredefined(currentCellNumberFormatId))
                {
                    if (currentCellFormatNode.NumberFormatId == NumberFormat.PercentageDoubleNumberFormatID
                        || currentCellFormatNode.NumberFormatId == NumberFormat.PercentageIntNumberFormatID)
                    {
                        return Convert.ToString(NumberFormat.Percentage);
                    }
                    else
                    {
                        return "General";
                    }
                }
                else
                {
                    return Convert.ToString(currentCellNumberFormatId.FormatCode);
                }
            }
        }

        private static bool IsFormatPredefined(NumberingFormat currentCellNumberFormatId)
        {
            //A list of predefined values that aren't shown in XML can be found at:
            //https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
            return currentCellNumberFormatId == null;
        }

        private static void GetCurrentCellNumberFormatIdNode(Stylesheet currentDocumentStylesheet,
                                CellFormat currentCellFormatNode, out NumberingFormat currentCellNumberFormatId)
        {
            currentCellNumberFormatId = currentDocumentStylesheet.NumberingFormats?.Descendants<NumberingFormat>()
            .Where(NumberFormatElementsFromSheet =>
            NumberFormatElementsFromSheet.NumberFormatId.Value == currentCellFormatNode.NumberFormatId).FirstOrDefault();
        }

        private static void GetCurrentCellFormatNode(Cell cell, Stylesheet currentDocumentStylesheet, out CellFormat currentCellFormatNode)
        {
            currentCellFormatNode = (CellFormat)currentDocumentStylesheet.CellFormats.ElementAt(int.Parse(cell.StyleIndex.Value.ToString()));
        }

        private static void GetStyleSheet(SpreadsheetDocument document, out Stylesheet currentDocumentStylesheet)
        {
            currentDocumentStylesheet = CommonExcelObjects.GetWorkbookStyleSheet(document);
        }

        private static bool CellDoesNotHaveStyleIndexAttribute(Cell cell)
        {
            return cell.StyleIndex == null;
        }
    }

    internal static class CellComparator
    {
        public static bool CompareCellsXML(object expected, object actual, string expectedFormat)
        {
            if (double.TryParse(expected.ToString(), out double expectedValue) 
                && double.TryParse(actual.ToString(), out double actualValue))
            {
                return Math.Abs(expectedValue - actualValue) <= GetThreshold(expectedFormat);
            }
            return expected.Equals(actual);
        }

        public static bool CompareCells(double expected, double actual, string expectedFormat)
        {
            double threshold = GetThreshold(expectedFormat);
            return Math.Abs(expected - actual) <= threshold;
        }

        public static bool CompareCells(object expected, object actual, string expectedFormat)
        {
            return expected.Equals(actual);
        }

        private static double GetThreshold(string format)
        {
            switch (format)
            {
                case NumberFormat.Percentage:
                    return Settings.Threshold.Percentage;
                default:
                    return Settings.Threshold.Currency;
            }
        }
    }

    internal static class CellFormatter
    {
        public static string ApplyFormat(double value, string format)
        {
            switch (format)
            {
                case NumberFormat.Percentage:
                    return value.ToString(NumberFormat.Percentage);
                default:
                    return value.ToString();
            }
        }
        public static string ApplyFormat(object value, string format)
        {
            return (value != null)? Convert.ToString(value) : null;
        }

        public static string XmlApplyFormat(dynamic value, string format)
        {
            if (value != null)
            {
                Regex regex = new Regex(@"^[0-9]+\.[0-9]+");
                var match = regex.Match(Convert.ToString(value));

                if (match.Success && format == NumberFormat.Percentage)
                {
                    return Convert.ToDecimal(value).ToString(NumberFormat.Percentage);
                }
                else if (match.Success && format.Contains("0.000"))
                {
                    return Convert.ToDecimal(value).ToString("0.000");
                }
                else
                {
                    return value.ToString();
                }
            }
            else
            {
                return null;
            }
        }
    }

    internal static class NumberFormat
    {
        public const string Percentage = "0.00%";
        public const int PercentageDoubleNumberFormatID = 10;
        public const int PercentageIntNumberFormatID = 9;
    }

    public class CellComment
    {
        public int row;
        public int col;
        public string text;
        public CellComment(int i, int j, string comment)
        {
            row = i;
            col = j;
            text = comment;
        }
    }
}
