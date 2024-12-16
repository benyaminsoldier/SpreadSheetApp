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
            
            for (int i = 0; i < cols; i++) // teste Patricia
            {
                Column = new DataColumn();

                if (i < 26) // Colunas A-Z
                {
                    Column.ColumnName = $"{(char)(i + 65)}"; // 'A' é 65 no ASCII
                }
                else // Colunas além de Z (AA, AB, etc.)
                {
                    int quotient = (i) / 26 - 1; // Calcula o prefixo para letras duplas
                    int remainder = (i) % 26;   // Calcula o sufixo
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
        public DataSource(SpreadsheetDocument openedFile, int numOfRows, int numOfCols)
        {
            // firt we'll be creating the columns
            DataColumn Column;
            string columnName = "";

            for (int i = 0; i < numOfCols; i++) // teste Patricia
            {
                Column = new DataColumn();

                if (i < 26) // Colunas A-Z
                {
                    Column.ColumnName = $"{(char)(i + 65)}"; // 'A' é 65 no ASCII
                }
                else // Colunas além de Z (AA, AB, etc.)
                {
                    int quotient = (i) / 26 - 1; // Calcula o prefixo para letras duplas
                    int remainder = (i) % 26;   // Calcula o sufixo
                    Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
                }

                Column.DataType = typeof(string);
                Column.AllowDBNull = true;
                Column.DefaultValue = "";
                Column.MaxLength = 255;
                this.Columns.Add(Column);
            }

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
