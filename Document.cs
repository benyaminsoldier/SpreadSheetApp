using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace spreadsheetApp
{
    public partial class Document : Form
    {
        public string FileName { get; set; } = "";
        public int NumOfRows { get; set; }
        public int NumOfColumns { get; set; }
        public string FilePath { get; set; }
        public DateTime OriginDate { get; set; }
        public DateTime LastModificationDate { get; set; }
        public List<DataTable> DataTables { get; set; }
        public List<DataGridView> Layouts { get; set; }
        public DataGridView CurrentLayout { get; set; }
        public DataTable CurrentDataTable { get; set; }

        public Document(string name, int numOfRows, int numOfColumns, string filePath)
        {
            FileName = name;
            NumOfRows = numOfRows;
            NumOfColumns = numOfColumns;
            FilePath = filePath;
            OriginDate = DateTime.Now;
            LastModificationDate = DateTime.Now;
            CurrentDataTable = new DataSource(numOfRows, numOfColumns);
            CurrentLayout = new Sheet(CurrentDataTable);
            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }
        public Document(string name, int numOfRows, int numOfColumns, string filePath, SpreadsheetDocument fileToBeOpened)
        {
            FileName = name;
            NumOfRows = numOfRows;
            NumOfColumns = numOfColumns;
            FilePath = filePath;
            OriginDate = DateTime.Now;
            LastModificationDate = DateTime.Now;
            CurrentDataTable = new DataSource(fileToBeOpened, numOfRows, numOfColumns);
            //CurrentDataTable = CurrentDataTable.TransferDataToTable(fileToBeOpened, numOfRows, numOfColumns);
            CurrentLayout = new Sheet(CurrentDataTable);
            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }



        // ---------------------------------------------- DATA LOGIC BUSINESS LOGIC DATATABLE VIRTUAL SHEET ----------------------------------------

        //private DataTable TransferDataToTable(SpreadsheetDocument openedFile)
        //{
        //    DataTable tableToFill = new DataTable();
        //    tableToFill = AddColumnHeaderToTable(tableToFill);

        //    WorkbookPart workbookPart = openedFile.WorkbookPart;
        //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();
        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

        //    var rows = worksheetPart.Worksheet.Descendants<Row>();

        //    foreach (Row row in rows)
        //    {
        //        DataRow dataRow = tableToFill.NewRow();
        //        int columnIndex = 0;

        //        foreach (Cell cell in row.Descendants<Cell>())
        //        {
        //            string cellValue = GetCellValue(workbookPart, cell);
        //            if (columnIndex < tableToFill.Columns.Count)
        //            {
        //                dataRow[columnIndex] = cellValue;
        //            }
        //            columnIndex++;
        //        }
        //        tableToFill.Rows.Add(dataRow);
        //    }
        //    return tableToFill;
        //}

        //private DataTable AddColumnHeaderToTable(DataTable table)
        //{
        //    DataColumn Column;
        //    string columnName = "";
        //    for (int i = 1; i < NumOfColumns; i++)
        //    {
        //        if (i <= 26)
        //        {
        //            // First 26 columns are just A-Z.
        //            columnName = $"{(char)(i + 64)}"; // 'A' is 65 in ASCII, so adding 64 to get A-Z.
        //            Column = new DataColumn(columnName);
        //        }
        //        else
        //        {
        //            // For columns beyond Z (i.e., AA, AB, etc.)
        //            int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
        //            int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
        //            // Combine the prefix and suffix to get AA, AB, etc.
        //            columnName = $"{(char)(quotient + 64)}{(char)(remainder + 64)}";
        //            Column = new DataColumn(columnName);
        //        }
        //        Column.DataType = typeof(string);
        //        Column.AllowDBNull = true;
        //        Column.DefaultValue = "";
        //        Column.MaxLength = 255;
        //        table.Columns.Add(Column);
        //    }
        //    return table;
        //}
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
        private void DisplayLayout(DataGridView sheet)
        {
            Controls.Add(sheet);
            sheet.Focus();
        }
        public void Display()
        {
            this.Show();
        }

        private void Document_Load(object sender, EventArgs e)
        {
            colorsPallette2.ChosenColor += (s, e) => { /*SheetCell Method pending*/};
        }
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.FileName = "";
                //sfd.Filter = "Excel|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // USING OPENXML
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Create(sfd.FileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = doc.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                        sheets.Append(sheet);
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // Write DataTable rows
                        foreach (DataRow dataRow in this.CurrentDataTable.Rows)
                        {
                            Row newRow = new Row();
                            foreach (var item in dataRow.ItemArray)
                            {
                                Cell cell = new Cell
                                {
                                    DataType = CellValues.String,
                                    CellValue = new CellValue(item.ToString())
                                };
                                newRow.AppendChild(cell);
                            }
                            sheetData.AppendChild(newRow);
                        }
                        workbookPart.Workbook.Save();
                    }
                }
            }
        }

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {
            colorsPallette2.BringToFront();
        }

        public class DocParams
        {
            private string title;
            private int rows;
            private int columns;
            public string Title { get => title; set => title = validateTitle(value); }
            public int Rows { get => rows; set => rows = validateRows(value); }
            public int Columns { get => columns; set => columns = validateCols(value); }
            public DocParams()
            {
                title = "Document1";
                rows = 12;
                columns = 12;
            }
            private string validateTitle(string title)
            {
                if (string.IsNullOrEmpty(title)) throw new Exception("Invalid document's title");
                return title;
            }
            private int validateRows(int rows)
            {
                if (int.IsPositive(rows)) return rows;
                else throw new Exception("Invalid number of rows");

            }
            private int validateCols(int cols)
            {
                if (int.IsPositive(cols)) return cols;
                else throw new Exception("Invalid number of columns");

            }
        }

        private void toolStripSplitButton1_Click(object sender, EventArgs e)
        {
            
            if (colorsPallette2.Visible) colorsPallette2.SendToBack();
            else colorsPallette2.BringToFront();

            colorsPallette2.Visible = !colorsPallette2.Visible;
        }
    }
}
