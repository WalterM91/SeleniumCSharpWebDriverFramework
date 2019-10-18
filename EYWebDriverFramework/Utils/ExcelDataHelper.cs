using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace EYWebDriverFramework.Utils
{
    public static class ExcelComparator
    {
        public static Exception ex = null;
        internal static string GetCellName(int row, int col)
        {
            int dividend = col;
            string columnName = String.Empty;
            int module;

            while (dividend > 0)
            {
                module = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + module).ToString() + columnName;
                dividend = (int)((dividend - module) / 26);
            }

            return columnName + row.ToString();
        }

        private static void CompareFormatError(ref Dictionary<string, List<Tuple<int, int>>> cellsErrors, DataTable expectedTable, DataTable actualTable)
        {
            // Get the min Row and Col number so the comparison is between all
            int minRow = Math.Min(expectedTable.Rows.Count,actualTable.Rows.Count);
            int minCol = Math.Min(expectedTable.Columns.Count, actualTable.Columns.Count);
            int error = 0;
            cellsErrors.Add(expectedTable.TableName, new List<Tuple<int, int>>());
            for (int i = 0; i < minRow; i++)
            {
                for (int j = 0; j < minCol; j++)
                {
                    var expectedCell = expectedTable.Rows[i].ItemArray[j];
                    var actualCell = actualTable.Rows[i].ItemArray[j];
                    if (!expectedCell.Equals(actualCell))
                    {
                        cellsErrors[expectedTable.TableName].Add(new Tuple<int, int>(i + 1, j + 1));
                        error++;
                    }
                }
                if (i >= Constants.MAX_FORMAT_ROW && error > Constants.MAX_FORMAT_ERRORS)
                    break;
            }
        }

        public static Dictionary<string, List<Tuple<int, int>>> CompareExcelDataTable(string expectedFile, string actualFile)
        {
            ex = null;
            Dictionary<string, List<Tuple<int, int>>> cellsErrors = new Dictionary<string, List<Tuple<int, int>>>();
            DataTableCollection actualTables = ExcelReader.ExcelToDataTable(actualFile);
            DataTableCollection expectedTables = ExcelReader.ExcelToDataTable(expectedFile);
            foreach (DataTable expectedTable in expectedTables)
            {
                int perc = 0, error = 0;
                int total = expectedTable.Rows.Count * expectedTable.Columns.Count;
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                DataTable actualTable = actualTables[expectedTable.TableName];
                // Missing tab error
                if (actualTable == null)
                {
                    cellsErrors.Add(expectedTable.TableName, new List<Tuple<int, int>>());
                    cellsErrors[expectedTable.TableName].Add(new Tuple<int, int>(1, 1));
                    ex = new FormatException("Discrepancies found in the format\nActual Report has missing tabs.");
                    continue;
                }

                // Format error
                if (actualTable.Rows.Count != expectedTable.Rows.Count || actualTable.Columns.Count != expectedTable.Columns.Count)
                {
                    ex = new FormatException(string.Format("Discrepancies found in the format\nExpected Report Rows: {0} Columns: {1} \nActual Report Rows: {2} Columns: {3}",
                        expectedTable.Rows.Count, expectedTable.Columns.Count,
                        actualTable.Rows.Count, actualTable.Columns.Count));
                    CompareFormatError(ref cellsErrors, expectedTable, actualTable);
                    continue;
                }

                cellsErrors.Add(expectedTable.TableName, new List<Tuple<int, int>>());
                for (int i = 0; i < expectedTable.Rows.Count; i++)
                {
                    for (int j = 0; j < expectedTable.Columns.Count; j++)
                    {
                        var expectedCell = expectedTable.Rows[i].ItemArray[j];
                        var actualCell = actualTable.Rows[i].ItemArray[j];
                        if (!expectedCell.Equals(actualCell))
                        {
                            cellsErrors[expectedTable.TableName].Add(new Tuple<int, int>(i + 1, j + 1));
                            error++;
                        }
                        perc++;
                    }
                }
                LogHelpers.Write("Completed: " + (perc * 100 / total) + "% " + perc + "/" + total + " Potential errors found:" + error + " Elapsed Time: " + stopWatch.Elapsed.TotalSeconds);
            }
            return cellsErrors;
        }
    }


    public static class ExcelInputReader<T>
    {
        private const int PROJECTLINK_HEADER = 1;
        private const int TESTCASE_HEADER = 2;

        private static readonly List<T> _dataCol = new List<T>();


        public static void PopulateInReportCollection(string fileName, Func<DataRow, string, string, string, T> loadObj, string tableName = "Tests")
        {
            DataTable table = ExcelReader.ExcelToDataTable(fileName)[tableName];
            var projectLink = "";
            var projectName = "";
            var setName = "";
            //Iterate through rows and collumns of table
            for (int row = 1; row < table.Rows.Count; row++)
            {
                switch (row)
                {
                    case PROJECTLINK_HEADER:
                        // First row for project link
                        projectLink = table.Rows[row][0].ToString();
                        projectName = table.Rows[row][1].ToString();
                        setName = table.Rows[row][2].ToString();
                        break;
                    case TESTCASE_HEADER:
                        // Test cases header
                        break;
                    default:
                        T dtTable = loadObj(table.Rows[row], projectLink, projectName, setName);
                        //Add all the details for each row
                        _dataCol.Add(dtTable);
                        break;
                }

            }
        }

        public static void PopulateInReviewCalcCollection(string fileName, Func<DataRow, T> loadObj, string tableName = "Tests")
        {
            DataTable table = ExcelReader.ExcelToDataTable(fileName)[tableName];
            //Iterate through rows and collumns of table
            for (int row = 0; row < table.Rows.Count; row++)
            {
                switch (row)
                {
                    case PROJECTLINK_HEADER:
                        // First row for project link
                        T dtTable = loadObj(table.Rows[row]);
                        //Add all the details for each row
                        _dataCol.Add(dtTable);
                        return;
                    default:
                        break;
                }

            }
        }

        public static List<T> GetData()
        {
            return _dataCol;
        }
    }

    public static class ExcelReader
    {
        public static DataTableCollection ExcelToDataTable(object filePathOrStream)
        {
            var ds = new DataSet();
            SpreadsheetDocument document;
            document = OpenExcelFile(filePathOrStream);

            WorkbookPart WorkbookPart = document.WorkbookPart;
            Sheets sheets = WorkbookPart.Workbook.Sheets;
            foreach (Sheet sheet in sheets){
                //Skip Hidden tabs.
                if (sheet.State != null && sheet.State != SheetStateValues.Visible)
                    continue;

                var dataTable = new DataTable{
                    TableName = sheet.Name
                };
                WorksheetPart WorksheetPart = WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                FillTableRows(document, dataTable, WorksheetPart.Worksheet.GetFirstChild<SheetData>().Descendants<Row>());
                
                ds.Tables.Add(dataTable);
            }
            document.Dispose();
            return ds.Tables;
        }

        private static void FillTableRows(SpreadsheetDocument document, DataTable dataTable, IEnumerable<Row> rows)
        {
            int lastRow = 1;
            foreach (Row row in rows){
                if (row.RowIndex != null){
                    while (lastRow < (int)row.RowIndex.Value){
                        dataTable.Rows.Add();
                        lastRow++;
                        FillEmptyRow(dataTable, row);
                    }
                }
                lastRow++;
                dataTable.Rows.Add();
                FillRowWithValues(document, dataTable, row);
            }
            CleanEmptyRows(dataTable);
        }

        private static void CleanEmptyRows(DataTable dataTable){
            // Clean empty rows at the bottom
            while (EmptyRow(dataTable.Rows[dataTable.Rows.Count - 1]))
                dataTable.Rows.RemoveAt(dataTable.Rows.Count - 1);
        }

        private static SpreadsheetDocument OpenExcelFile(object filePathOrStream){
            SpreadsheetDocument document;
            if (filePathOrStream is Stream)
                document = SpreadsheetDocument.Open(filePathOrStream as Stream, false);
            else if (!string.IsNullOrEmpty(filePathOrStream?.ToString())){
                string path = filePathOrStream.ToString();
                document = SpreadsheetDocument.Open(Path.GetFullPath(path), false);
            }
            else
                throw new ArgumentException("Argument must be Stream or Path string.");
            return document;
        }

        private static void FillEmptyRow(DataTable dataTable, Row row){
            int i = 0;
            foreach (Cell cell in GetRowCells(row)){
                if (i == dataTable.Columns.Count)
                    dataTable.Columns.Add(i.ToString());
                dataTable.Rows[dataTable.Rows.Count - 1][i] = null;
                i++;
            }
        }

        private static void FillRowWithValues(SpreadsheetDocument document, DataTable dataTable, Row row)
        {
            int i = 0;
            foreach (Cell cell in GetRowCells(row)){
                if (i == dataTable.Columns.Count)
                    dataTable.Columns.Add(i.ToString());
                dataTable.Rows[dataTable.Rows.Count - 1][i] = GetCellValue(document, cell);
                i++;
            }
        }

        private static bool EmptyRow(DataRow row)
        {
            return (from item in row.ItemArray where string.IsNullOrEmpty(item.ToString()) select item).Count() == row.ItemArray.Length;
        }
               
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = string.Empty;
            // If the cell does not exist, return an empty string:
            if (cell != null){
                value = cell.InnerText;
                if (cell.DataType != null){
                    switch (cell.DataType.Value){
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared 
                            // strings table.
                            var stringTable = document.WorkbookPart.
                              GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something is 
                            // wrong. Return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null){
                                value = stringTable.SharedStringTable.
                                  ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value){
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
                else
                {
                    if (cell.InnerText.Contains("E")){ //número en notación científica
                        int test = int.Parse(cell.InnerText.Split('E').LastOrDefault());
                        double.TryParse(cell.InnerText, out double doubleValue);
                        value = doubleValue.ToString("0." + new string('#', Math.Abs(test) + doubleValue.ToString().Length));
                    }
                }
            }
            return value;
        }

        private static List<Cell> GetRowCells(Row row)
        {
            List<Cell> cells = new List<Cell>();

            int currentCount = 0;
            int xmlColumn = 0;

            foreach (Cell cell in row.Descendants<Cell>()){
                string columnName = GetColumnName(cell.CellReference); 

                int currentColumnIndex = (!string.IsNullOrEmpty(columnName)) ? ConvertColumnNameToNumberZeroBased(columnName) : currentCount;

                for (; currentCount < currentColumnIndex; currentCount++){
                    cells.Add(new Cell());
                }

                cells.Add(cell);
                currentCount++;
                xmlColumn++;
            }
            return cells;
        }

        /// <summary>
        /// Iterates over all row and cell nodes to check and add (if missing) RowIndex and CellReference attributes.
        /// Those values make it easier to set error style and add comments.
        /// </summary>
        /// <param name="document">File being modified.</param>
        /// <param name="worksheet">Current sheet that is going to be changed.</param>
        public static void StructurateExcelSheetDataXmlAttributes(SpreadsheetDocument document, Worksheet worksheet)
        {
            //Row and Cell nodes are contained within a sheetData node, there is one sheetData node per worksheet.

            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            bool firstRow = true;
            int currentRow = 0;
            int currentRowIndex = 0;

            foreach (Row rowItem in sheetData){
                if (rowItem.RowIndex == null){
                    if (firstRow){
                        currentRow = 1;
                        currentRowIndex = 1;
                    }
                    else{
                        currentRow++;
                        currentRowIndex++;
                    }
                    rowItem.RowIndex = (uint)currentRowIndex;
                    firstRow = false;
                }
                else{
                    currentRow++;
                    firstRow = false;
                    currentRowIndex = (int)rowItem.RowIndex.Value;
                    while ((currentRow) < currentRowIndex){
                        sheetData.InsertBefore(new Row() { RowIndex = (uint)(currentRow) }, rowItem);
                        currentRow++;
                    }
                }
                StructurateCellsInRow(rowItem);
            }
            if (document.FileOpenAccess != FileAccess.Read){
                document.Save();
            }
        }

        private static void StructurateCellsInRow(Row rowItem)
        {
            bool firstCell = true;
            int currentCol = 0;
            int currentColIndex = 0;
            foreach (Cell cellItem in rowItem.Descendants<Cell>()){
                if (cellItem.CellReference == null){
                    if (firstCell){
                        currentCol = 1;
                        currentColIndex = 1;
                    }
                    else{
                        currentCol++;
                        currentColIndex++;
                    }
                    cellItem.CellReference = ConvertColumnNumberToName(currentColIndex) + rowItem.RowIndex.Value;
                    firstCell = false;
                }
                else{
                    currentCol++;
                    firstCell = false;
                    currentColIndex = ConvertColumnNameToNumberZeroBased(GetColumnName(cellItem.CellReference)) + 1;
                    while ((currentCol) < currentColIndex){
                        rowItem.InsertBefore(new Cell() { CellReference = ConvertColumnNumberToName(currentCol) + rowItem.RowIndex.Value }, cellItem);
                        currentCol++;
                    }
                }
            }
        }

        public static Cell GetCell(Worksheet worksheet, int rowIndex, int colIndex)
        {
            Row row = GetRow(worksheet, (uint)rowIndex);

            if (row == null)
                return new Cell();

            Cell cell = row.Elements<Cell>().Where(c => string.Compare
                      (c.CellReference.Value, ConvertColumnNumberToName(colIndex) +
                      rowIndex, true) == 0).FirstOrDefault();

            if (cell == null){
                return new Cell();
            }

            return cell;
        }

        // Given a worksheet and a row index, return the row.
        public static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
                  Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
        }
        
        private static string GetColumnName(string cellReference)
        {
            if (cellReference == null)
                return "";
            // Match the column name portion of the cell name.
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);

            return match.Value;
        }

        private static int ConvertColumnNameToNumberZeroBased(string columnName)
        {
            var alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            int convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++){
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64; // ASCII 'A' = 65
                convertedValue += current * (int)Math.Pow(26, i);
            }
            return convertedValue;
        }

        private static string ConvertColumnNumberToName(int index)
        {
            int columnId = index - 1;
            string cellColumn = string.Empty;
            if (columnId < 0) return cellColumn;
            int a = 26;
            int x = (int)Math.Floor(Math.Log((columnId) * (a - 1) / a + 1, a));
            columnId -= (int)(Math.Pow(a, x) - 1) * a / (a - 1);
            for (int i = x + 1; columnId + i > 0; i--){
                cellColumn = ((char)(65 + columnId % a)).ToString() + cellColumn;
                columnId /= a;
            }
            return cellColumn;
        }

        public static SpreadsheetDocument OpenExcelFile(object filePathOrStream, bool isEditable)
        {
            SpreadsheetDocument document;
            if (filePathOrStream is Stream)
                document = SpreadsheetDocument.Open(filePathOrStream as Stream, isEditable);
            else if (!string.IsNullOrEmpty(filePathOrStream?.ToString()))
            {
                string path = filePathOrStream.ToString();
                document = SpreadsheetDocument.Open(Path.GetFullPath(path), isEditable);
            }
            else
                throw new ArgumentException("Argument must be Stream or Path string.");
            return document;
        }

        /// <summary>
        /// Adds all the comments defined in the List to the current worksheet.
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part of file.</param>
        /// <param name="commentsToAddList">List of CellComment which contain cell coordinates and the text value to set as comment.</param>
        public static void InsertComments(WorksheetPart worksheetPart, List<CellComment> commentsToAddList)
        {
            if (commentsToAddList.Any())
            {
                string commentsVmlXml = string.Empty;

                // Create all the comment VML Shape XML
                foreach (var commentToAdd in commentsToAddList)
                {
                    commentsVmlXml += GetCommentVMLShapeXML(ConvertColumnNumberToName(commentToAdd.col), commentToAdd.row.ToString());
                }

                // The VMLDrawingPart should contain all the definitions for how to draw every comment shape for the worksheet
                VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
                using (XmlTextWriter writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8))
                {

                    writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n" +
                    "</o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n  path=\"m,l,21600r21600,l21600,xe\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n </v:shapetype>"
                    + commentsVmlXml + "</xml>");
                }

                // Create the comment elements
                foreach (var commentToAdd in commentsToAddList)
                {
                    WorksheetCommentsPart worksheetCommentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();

                    // We only want one legacy drawing element per worksheet for comments
                    if (worksheetPart.Worksheet.Descendants<LegacyDrawing>().SingleOrDefault() == null)
                    {
                        string vmlPartId = worksheetPart.GetIdOfPart(vmlDrawingPart);
                        LegacyDrawing legacyDrawing = new LegacyDrawing() { Id = vmlPartId };
                        worksheetPart.Worksheet.Append(legacyDrawing);
                    }

                    Comments comments;
                    bool appendComments = false;
                    if (worksheetPart.WorksheetCommentsPart.Comments != null)
                    {
                        comments = worksheetPart.WorksheetCommentsPart.Comments;
                    }
                    else
                    {
                        comments = new Comments();
                        appendComments = true;
                    }

                    // We only want one Author element per Comments element
                    if (worksheetPart.WorksheetCommentsPart.Comments == null)
                    {
                        Authors authors = new Authors();
                        Author author = new Author
                        {
                            Text = "Author Name"
                        };
                        authors.Append(author);
                        comments.Append(authors);
                    }

                    CommentList commentList;
                    bool appendCommentList = false;
                    if (worksheetPart.WorksheetCommentsPart.Comments != null &&
                        worksheetPart.WorksheetCommentsPart.Comments.Descendants<CommentList>().SingleOrDefault() != null)
                    {
                        commentList = worksheetPart.WorksheetCommentsPart.Comments.Descendants<CommentList>().Single();
                    }
                    else
                    {
                        commentList = new CommentList();
                        appendCommentList = true;
                    }
                    Comment comment = new Comment() { Reference = string.Concat(ConvertColumnNumberToName(commentToAdd.col), commentToAdd.row), AuthorId = (UInt32Value)0U };

                    CommentText commentTextElement = new CommentText();

                    Run run = new Run();

                    RunProperties runProperties = new RunProperties();
                    Bold bold = new Bold();
                    FontSize fontSize = new FontSize() { Val = 8D };
                    Color color = new Color() { Indexed = (UInt32Value)81U };
                    RunFont runFont = new RunFont() { Val = "Tahoma" };
                    RunPropertyCharSet runPropertyCharSet = new RunPropertyCharSet() { Val = 1 };

                    runProperties.Append(bold);
                    runProperties.Append(fontSize);
                    runProperties.Append(color);
                    runProperties.Append(runFont);
                    runProperties.Append(runPropertyCharSet);
                    Text text = new Text
                    {
                        Text = commentToAdd.text
                    };

                    run.Append(runProperties);
                    run.Append(text);

                    commentTextElement.Append(run);
                    comment.Append(commentTextElement);
                    commentList.Append(comment);

                    // Only append the Comment List if this is the first time adding a comment
                    if (appendCommentList)
                    {
                        comments.Append(commentList);
                    }

                    // Only append the Comments if this is the first time adding Comments
                    if (appendComments)
                    {
                        worksheetCommentsPart.Comments = comments;
                    }
                }
            }
        }

        /// <summary>
        /// Creates the VML Shape XML for a comment. It determines the positioning of the
        /// comment in the excel document based on the column name and row index.
        /// </summary>
        /// <param name="columnName">Column name containing the comment</param>
        /// <param name="rowIndex">Row index containing the comment</param>
        /// <returns>VML Shape XML for a comment</returns>
        private static string GetCommentVMLShapeXML(string columnName, string rowIndex)
        {
            string commentVmlXml = string.Empty;

            // Parse the row index into an int so we can subtract one
            int commentRowIndex;
            if (int.TryParse(rowIndex, out commentRowIndex))
            {
                commentRowIndex -= 1;

                commentVmlXml = "<v:shape id=\"" + Guid.NewGuid().ToString().Replace("-", "") + "\" type=\"#_x0000_t202\" style=\'position:absolute;\r\n  margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;\r\n  visibility:hidden\' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\r\n  <v:fill color2=\"#ffffe1\"/>\r\n" +
                "<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\r\n  <v:path o:connecttype=\"none\"/>\r\n  <v:textbox style=\'mso-fit-shape-to-text:true'>\r\n   <div style=\'text-align:left\'></div>\r\n  </v:textbox>\r\n  <x:ClientData ObjectType=\"Note\">\r\n   <x:MoveWithCells/>\r\n" +
                "<x:SizeWithCells/>\r\n   <x:Anchor>\r\n" + GetAnchorCoordinatesForVMLCommentShape(columnName, rowIndex) + "</x:Anchor>\r\n   <x:AutoFill>False</x:AutoFill>\r\n   <x:Row>" + commentRowIndex + "</x:Row>\r\n   <x:Column>" + ConvertColumnNameToNumberZeroBased(columnName) + "</x:Column>\r\n  </x:ClientData>\r\n </v:shape>";
            }

            return commentVmlXml;
        }

        /// <summary>
        /// Gets the coordinates for where on the excel spreadsheet to display the VML comment shape
        /// </summary>
        /// <param name="columnName">Column name of where the comment is located (ie. B)</param>
        /// <param name="rowIndex">Row index of where the comment is located (ie. 2)</param>
        /// <returns><see cref="<x:Anchor>"/> coordinates in the form of a comma separated list</returns>
        private static string GetAnchorCoordinatesForVMLCommentShape(string columnName, string rowIndex)
        {
            string coordinates = string.Empty;
            int startingRow;
            int startingColumn = ConvertColumnNameToNumberZeroBased(columnName);

            // From (upper right coordinate of a rectangle)
            // [0] Left column
            // [1] Left column offset
            // [2] Left row
            // [3] Left row offset
            // To (bottom right coordinate of a rectangle)
            // [4] Right column
            // [5] Right column offset
            // [6] Right row
            // [7] Right row offset
            List<int> coordList = new List<int>(8) { 0, 0, 0, 0, 0, 0, 0, 0 };

            if (int.TryParse(rowIndex, out startingRow))
            {
                // Make the row be a zero based index
                startingRow -= 1;

                coordList[0] = startingColumn + 1; // If starting column is A, display shape in column B
                coordList[1] = 15;
                coordList[2] = startingRow;
                coordList[4] = startingColumn + 3; // If starting column is A, display shape till column D
                coordList[5] = 15;
                coordList[6] = startingRow + 3; // If starting row is 0, display 3 rows down to row 3

                // The row offsets change if the shape is defined in the first row
                if (startingRow == 0)
                {
                    coordList[3] = 2;
                    coordList[7] = 16;
                }
                else
                {
                    coordList[3] = 10;
                    coordList[7] = 4;
                }

                coordinates = string.Join(",", coordList.ConvertAll<string>(x => x.ToString()).ToArray());
            }

            return coordinates;
        }




    }

    public static class StyleSheetFunctions
    {
        public static UInt32Value ErrorStyleFillId = 2;
        public static UInt32Value DefaultStyle = 0;

        public static uint GetOrSetErrorFillID(SpreadsheetDocument workdocument)
        {
            WorkbookStylesPart stylesPart = workdocument.WorkbookPart.WorkbookStylesPart;
            Fills fills = stylesPart.Stylesheet.Fills;
            Fill cellFill = null;
            uint fillId = 0;

            if (!ExistsRedFillPattern(fills, ref cellFill, ref fillId))
            {
                Fill newErrorFill = new Fill();
                PatternFill newErrorPattern = new PatternFill() { PatternType = PatternValues.Solid };
                newErrorPattern.AppendChild(new ForegroundColor() { Rgb = "FFFF0000" });
                newErrorPattern.AppendChild(new BackgroundColor() { Indexed = 64 });
                newErrorFill.AppendChild(newErrorPattern);
                fills.AppendChild(newErrorFill);
                fills.Count++;
            }
            workdocument.Save();
            ErrorStyleFillId = fillId;
            return fillId;
        }

        private static bool ExistsRedFillPattern(Fills fills, ref Fill cellFill, ref uint fillId)
        {
            foreach (Fill fillItem in fills)
            {
                ForegroundColor foreItem = fillItem.Descendants<PatternFill>().FirstOrDefault().Descendants<ForegroundColor>().FirstOrDefault();
                if (foreItem != null)
                {
                    if (foreItem.Rgb == "FFFF0000")
                    {
                        cellFill = fillItem;
                        break;
                    }
                }
                fillId++;
            }
            return cellFill != null;
        }

        public static void GetAndOrSetErrorStyleID(SpreadsheetDocument workdocument, Worksheet worksheet, Cell currentCell)
        {
            WorkbookStylesPart stylesPart = workdocument.WorkbookPart.WorkbookStylesPart;
            uint fillId = GetOrSetErrorFillID(workdocument);
            uint styleId = currentCell.StyleIndex ?? 0;
            CellFormats cellFormats = stylesPart.Stylesheet.CellFormats;
            CellFormat currentCellCellFormat = cellFormats.Descendants<CellFormat>().ElementAt((int)styleId);
            bool cellFormatComparation = true;
            uint checkedStyleIndex = 0;
            //iterate over all cellFormat of the page
            foreach (CellFormat cfItem in cellFormats)
            {
                checkedStyleIndex++;

                cellFormatComparation = true;
                //iterate over all attributes of the current cellFormat to check if all are the same
                foreach (var item in cfItem.GetAttributes()) 
                {
                    bool check = false;
                    try
                    {
                        check = Equals(currentCellCellFormat.GetAttribute(item.LocalName, item.NamespaceUri), item);
                    }
                    catch (KeyNotFoundException)
                    {
                        check = false;
                        break;
                    }
                    if (!check) //
                    {
                        cellFormatComparation = false;
                        break;
                    }
                }
                if (cellFormatComparation)
                {
                    styleId = checkedStyleIndex;
                    break;
                }
            }
            if (cellFormatComparation)
            {
                if (currentCellCellFormat.FillId != fillId)
                {
                    CellFormat currentCellNewFormat = (CellFormat)currentCellCellFormat.CloneNode(true);
                    currentCellNewFormat.FillId = fillId;
                    cellFormats.AppendChild(currentCellNewFormat);
                    cellFormats.Count++;
                }
            }
            else
            {
                cellFormats.AppendChild(new CellFormat()
                    { BorderId = DefaultStyle, FillId = ErrorStyleFillId, FontId = DefaultStyle, NumberFormatId = DefaultStyle });
                cellFormats.Count++;
            }
            styleId = (uint)cellFormats.Descendants<CellFormat>().Count();
            currentCell.StyleIndex = styleId - 1;
            worksheet.Save();
        }
    }

    public static class CommonExcelObjects
    {
        public static WorkbookStylesPart GetWorkbookStylesPart(SpreadsheetDocument spreadsheetDocument) =>
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart;
        public static Stylesheet GetWorkbookStyleSheet(SpreadsheetDocument spreadsheetDocument) =>
            GetWorkbookStylesPart(spreadsheetDocument).Stylesheet;
    }

}