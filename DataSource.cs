using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    internal class DataSource : DataTable
    {
        public DataSource(int rows, int cols) 
        {
            DataColumn Column;
            string columnName = "";
            for (int i = 0; i < cols; i++)
            {
                if (i <= 26) {
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(i + 65)}"; 
                } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
                else
                {
                    // For columns beyond Z (i.e., AA, AB, etc.)
                    int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
                    int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
                    // Combine the prefix and suffix to get AA, AB, etc.
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
                }
                Column.DataType = typeof(string);
                Column.AllowDBNull = true;
                Column.DefaultValue = "";
                Column.MaxLength = 255;
                this.Columns.Add(Column);
            }
            for (int j = 0; j < rows; j++)
            {
                this.Rows.Add(this.NewRow());
            }
        }
        //public DataSource(SpreadsheetDocument openedFile, int numOfRows, int numOfCols)
        //{
        //    DataColumn Column;
        //    string columnName = "";
        //    for (int i = 0; i < numOfCols; i++)
        //    {
        //        if (i <= 26)
        //        {
        //            Column = new DataColumn();
        //            Column.ColumnName = $"{(char)(i + 65)}";
        //        } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
        //        else
        //        {
        //            // For columns beyond Z (i.e., AA, AB, etc.)
        //            int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
        //            int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
        //            // Combine the prefix and suffix to get AA, AB, etc.
        //            Column = new DataColumn();
        //            Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
        //        }
        //        Column.DataType = typeof(string);
        //        Column.AllowDBNull = true;
        //        Column.DefaultValue = "";
        //        Column.MaxLength = 255;
        //        this.Columns.Add(Column);
        //    }

        //    WorkbookPart workbookPart = openedFile.WorkbookPart;
        //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();
        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

        //    var rows = worksheetPart.Worksheet.Descendants<Row>();

        //    //int rowIndex = 0;
        //    foreach (Row row in rows)
        //    {
        //        DataRow dataRow = this.NewRow();
        //        int columnIndex = 0;

        //        foreach (Cell cell in row.Descendants<Cell>())
        //        {
        //            string cellValue = GetCellValue(workbookPart, cell);
        //            if (columnIndex < this.Columns.Count)
        //            {
        //                dataRow[columnIndex] = cellValue;
        //                //this.Rows.Add(dataRow);
        //            }
        //            columnIndex++;

        //        }
        //        this.Rows.Add(dataRow);
        //        //this.Rows[rowIndex].Add(dataRow);
        //    }
        //}
        //private string GetCellValue(WorkbookPart workbookPart, Cell cell)
        //{
        //    if (cell == null || cell.CellValue == null)
        //        return string.Empty;

        //    // If the cell contains a shared string, retrieve the value from the shared string table
        //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //    {
        //        var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
        //        return sharedStringTable.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText;
        //    }
        //    else
        //    {
        //        return cell.CellValue.InnerText;
        //    }
        //}
        public DataSource(SpreadsheetDocument openedFile, int numOfRows, int numOfCols)
        {
            // firt we'll be creating the columns
            DataColumn Column;
            string columnName = "";
            for (int i = 0; i < numOfCols; i++)
            {
                if (i <= 26)
                {
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(i + 65)}";
                } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
                else
                {
                    // For columns beyond Z (i.e., AA, AB, etc.)
                    int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
                    int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
                    // Combine the prefix and suffix to get AA, AB, etc.
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
                }
                Column.DataType = typeof(string);
                Column.AllowDBNull = true;
                Column.DefaultValue = "";
                Column.MaxLength = 255;
                this.Columns.Add(Column);
            }

            // retrieving data from the Spreadsheet document that we wish to open
            WorkbookPart workbookPart = openedFile.WorkbookPart;
            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            // acessing the first sheet and saving its rows in a variable
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var rows = sheetData.Elements<Row>();

            int rowIndex = 0;
            foreach (Row row in rows) // for each row of the worksheet
            {
                DataRow dataRow = this.NewRow();
                dataRow.BeginEdit();
                int columnIndex = 0;

                foreach (Cell cell in row.Descendants<Cell>())
                {
                    string cellValue = GetCellValue(cell, workbookPart);
                    
                    if (columnIndex < this.Columns.Count)
                    {
                        dataRow[columnIndex] = cellValue.ToString();
                        //this.Rows.Add(dataRow);
                    }
                    columnIndex++;
                }
                dataRow.EndEdit();
                this.Rows.Add(dataRow); 
            }
        }
        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string value = cell.CellValue?.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Handle shared strings
                SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                return sharedStringPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
