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
            //for (int i = 0; i < cols; i++) // comentado para Teste
            //{
            //    if (i <= 26) {
            //        Column = new DataColumn();
            //        Column.ColumnName = $"{(char)(i + 65)}"; 
            //    } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
            //    else
            //    {
            //        // For columns beyond Z (i.e., AA, AB, etc.)
            //        int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
            //        int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
            //        // Combine the prefix and suffix to get AA, AB, etc.
            //        Column = new DataColumn();
            //        Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
            //    }
            //    Column.DataType = typeof(string);
            //    Column.AllowDBNull = true;
            //    Column.DefaultValue = "";
            //    Column.MaxLength = 255;
            //    this.Columns.Add(Column);
            //}
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
            DataColumn Column;
            string columnName = "";
            //for (int i = 0; i < numOfCols; i++) //comentado para Testes Patricia
            //{
            //    if (i <= 26)
            //    {
            //        Column = new DataColumn();
            //        Column.ColumnName = $"{(char)(i + 65)}";
            //    } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
            //    else
            //    {
            //        // For columns beyond Z (i.e., AA, AB, etc.)
            //        int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
            //        int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
            //        // Combine the prefix and suffix to get AA, AB, etc.
            //        Column = new DataColumn();
            //        Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";
            //    }
            //    Column.DataType = typeof(string);
            //    Column.AllowDBNull = true;
            //    Column.DefaultValue = "";
            //    Column.MaxLength = 255;
            //    this.Columns.Add(Column);
            //}
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
            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            var rows = worksheetPart.Worksheet.Descendants<Row>();

            //int rowIndex = 0;
            //foreach (Row row in rows) // comentado por Patricia para testes
            //{
            //    DataRow dataRow = this.NewRow();
            //    int columnIndex = 0;

            //    foreach (Cell cell in row.Descendants<Cell>())
            //    {
            //        string cellValue = GetCellValue(workbookPart, cell);
            //        if (columnIndex < this.Columns.Count)
            //        {
            //            dataRow[columnIndex] = cellValue;
            //            //this.Rows.Add(dataRow);
            //        }
            //        columnIndex++;

            //    }
            //    this.Rows.Add(dataRow);
            //    //this.Rows[rowIndex].Add(dataRow);
            //}
            foreach (Row row in rows) // Patricia Testes
            {
                DataRow dataRow = this.NewRow();
                int columnIndex = 0;

                foreach (Cell cell in row.Descendants<Cell>())
                {
                    string cellValue = GetCellValue(workbookPart, cell);

                    // Preenche somente se o índice for válido
                    if (columnIndex < this.Columns.Count)
                    {
                        dataRow[columnIndex] = cellValue;
                    }

                    columnIndex++;
                }

                // Garante que todas as colunas sejam preenchidas
                for (int col = columnIndex; col < this.Columns.Count; col++)
                {
                    dataRow[col] = ""; // Preenche células vazias
                }

                this.Rows.Add(dataRow);
            }


        }
        private string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            // If the cell contains a shared string, retrieve the value from the shared string table
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                return sharedStringTable.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText;
            }
            else
            {
                return cell.CellValue.InnerText;
            }
        }
        //public DataTable TransferDataToTable(SpreadsheetDocument openedFile, int numOfRows, int numOfCols)
        //{
        //    DataSource dataSource = new DataSource(numOfRows, numOfCols);
        //    //DataTable tableToFill = new DataTable();
        //    //tableToFill = AddColumnHeaderToTable(tableToFill);

        //    WorkbookPart workbookPart = openedFile.WorkbookPart;
        //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();
        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

        //    var rows = worksheetPart.Worksheet.Descendants<Row>();

        //    foreach (Row row in rows)
        //    {
        //        DataRow dataRow = dataSource.NewRow();
        //        int columnIndex = 0;

        //        foreach (Cell cell in row.Descendants<Cell>())
        //        {
        //            string cellValue = GetCellValue(workbookPart, cell);
        //            if (columnIndex < dataSource.Columns.Count)
        //            {
        //                dataRow[columnIndex] = cellValue;
        //            }
        //            columnIndex++;
        //        }
        //        dataSource.Rows.Add(dataRow);
        //    }
        //    return dataSource;
        //}






    }
}
