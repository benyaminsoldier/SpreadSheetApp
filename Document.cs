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
            //DataTable emptyTable = new DataSource(numOfRows, numOfColumns);
            //CurrentDataTable = TransferDataToTable(emptyTable, fileToBeOpened);
            CurrentDataTable = new DataSource(fileToBeOpened, numOfRows, numOfColumns);

            //// debugging
            Console.WriteLine($"Rows: {CurrentDataTable.Rows.Count}, Columns: {CurrentDataTable.Columns.Count}");
            foreach (DataRow row in CurrentDataTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write($"{item}  ");
                }
                Console.WriteLine();
            }

            CurrentLayout = new Sheet(CurrentDataTable);
            CurrentLayout.DataSource = CurrentDataTable;
            CurrentLayout.Refresh(); // Ensure the UI is updated
            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }

        // ---------------------------------------------- DATA LOGIC BUSINESS LOGIC DATATABLE VIRTUAL SHEET ----------------------------------------

        //public DataTable TransferDataToTable(DataTable emptyTable, SpreadsheetDocument openedFile)
        //{
        //    WorkbookPart workbookPart = openedFile.WorkbookPart;
        //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();
        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

        //    var rows = worksheetPart.Worksheet.Descendants<Row>();

        //    foreach (Row row in rows)
        //    {
        //        DataRow dataRow = emptyTable.NewRow();
        //        int columnIndex = 0;

        //        foreach (Cell cell in row.Descendants<Cell>())
        //        {
        //            string cellValue = GetCellValue(cell, workbookPart);
        //            if (columnIndex < emptyTable.Columns.Count)
        //            {
        //                dataRow[columnIndex] = cellValue;
        //            }
        //            columnIndex++;
        //        }
        //        emptyTable.Rows.Add(dataRow);
        //    }
        //    return emptyTable;
        //}

        //private DataTable TransferDataToTable(DataTable initialTable, SpreadsheetDocument openedFile)
        //{
        //    // retrieving data from the Spreadsheet document that we wish to open
        //    WorkbookPart workbookPart = openedFile.WorkbookPart;
        //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        //    Worksheet worksheet = worksheetPart.Worksheet;

        //    // acessing the first sheet and saving its rows in a variable
        //    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        //    var rows = sheetData.Elements<Row>();

        //    int rowIndex = 0;
        //    foreach (Row row in rows) // for each row of the worksheet
        //    {
        //        DataRow dataRow = initialTable.NewRow();
        //        int columnIndex = 0;

        //        foreach (Cell cell in row.Descendants<Cell>())
        //        {
        //            string cellValue = GetCellValue(cell, workbookPart);
        //            if (columnIndex < initialTable.Columns.Count)
        //            {
        //                dataRow[columnIndex] = cellValue;
        //            }
        //            columnIndex++;
        //        }
        //        initialTable.Rows.Add(dataRow);
        //    }
        //    return initialTable;
        //}
        //private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        //{
        //    string value = cell.CellValue?.Text;
        //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //    {
        //        // Handle shared strings
        //        SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        //        return sharedStringPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        //    }
        //    return value;
        //}
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
            colorsPallette1.ChosenColor += (s, e) =>
            {
                foreach (var cell in CurrentLayout.SelectedCells)
                {
                    if (cell is SheetCell selectedCell)
                    {
                        selectedCell.ForeColor = e.ChosenColor;
                        CurrentLayout.InvalidateCell(selectedCell);
                    }
                }
            };
            colorsPallette2.ChosenColor += (s, e) =>
            {
                foreach (var cell in CurrentLayout.SelectedCells)
                {
                    if (cell is SheetCell selectedCell)
                    {
                        selectedCell.BackGroundColor = e.ChosenColor;
                        CurrentLayout.InvalidateCell(selectedCell);
                    }
                }
            };
            foreach (System.Windows.Forms.Control control in this.Controls)
            {
                if (!(control is ColorsPallette))
                {
                    control.Click += (s, e) =>
                    {
                        colorsPallette1.Visible = false;
                        colorsPallette2.Visible = false;
                    };
                }
            }
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

        private void BackGroundCellFormatBtn_Click(object sender, EventArgs e)
        {

            if (colorsPallette2.Visible) colorsPallette2.SendToBack();
            else colorsPallette2.BringToFront();

            colorsPallette2.Visible = !colorsPallette2.Visible;
        }

        private void btn_fontColor_Click(object sender, EventArgs e)
        {
            if (colorsPallette1.Visible) colorsPallette1.SendToBack();
            else colorsPallette1.BringToFront();

            colorsPallette1.Visible = !colorsPallette1.Visible;
        }

        private void btn_leftAlign_Click(object sender, EventArgs e)
        {
            foreach (var cell in CurrentLayout.SelectedCells)
            {
                if (cell is SheetCell selectedCell)
                {
                    selectedCell.Alignment = System.Drawing.StringAlignment.Near;
                    CurrentLayout.InvalidateCell(selectedCell);
                }
            }
        }

        private void btn_centerAlign_Click(object sender, EventArgs e)
        {
            foreach (var cell in CurrentLayout.SelectedCells)
            {
                if (cell is SheetCell selectedCell)
                {
                    selectedCell.Alignment = System.Drawing.StringAlignment.Center;
                    CurrentLayout.InvalidateCell(selectedCell);
                }
            }
        }

        private void btn_rightAlign_Click(object sender, EventArgs e)
        {
            foreach (var cell in CurrentLayout.SelectedCells)
            {
                if (cell is SheetCell selectedCell)
                {
                    selectedCell.Alignment = System.Drawing.StringAlignment.Far;
                    CurrentLayout.InvalidateCell(selectedCell);
                }
            }
        }

        private void avgMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void sumMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
