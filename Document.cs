using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static spreadsheetApp.SheetCell;

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
            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }
        // ---------------------------------------------- DATA LOGIC BUSINESS LOGIC DATATABLE VIRTUAL SHEET ----------------------------------------

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
                sfd.Title = "Save File As";
                sfd.FileName = "Untitled";
                sfd.Filter = "Excel Files(*.xlsx)| *.xlsx";
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

        private void pasteBtn_Click(object sender, EventArgs e)
        {
            IDataObject clipboardDataObject = Clipboard.GetDataObject();

            if (clipboardDataObject != null)
            {
                if (clipboardDataObject.GetDataPresent("SheetCellData"))
                {
                    SheetCell.DataCell cellData = clipboardDataObject.GetData("SheetCellData") as SheetCell.DataCell;

                    if (cellData != null)
                    {
                        SheetCell selectedCell = CurrentLayout.SelectedCells[0] as SheetCell;
                        selectedCell.Value = cellData.CellValue;
                        selectedCell.BackGroundColor = cellData.BackGroundColor;
                        selectedCell.ForeColor = cellData.ForeColor;
                        selectedCell.Alignment = cellData.Alignment;

                        CurrentLayout.InvalidateCell(selectedCell);

                    }
                    else
                    {
                        MessageBox.Show("Retrieved object is null. Check serialization1.");
                    }
                }
                else if(clipboardDataObject.GetDataPresent("SheetCellDataSet"))
                {

                    SheetCell.DataCellSet cellDataSet = clipboardDataObject.GetData("SheetCellDataSet") as SheetCell.DataCellSet;
                    if (cellDataSet != null)
                    {
                        List<SheetCell> cellsToBePasted = new List<SheetCell>();

                        SheetCell selectedCell = CurrentLayout.SelectedCells[0] as SheetCell;

                        int selectedRow = selectedCell.RowIndex;
                        int selectedColumn = selectedCell.ColumnIndex;

                        int lastRow  = selectedCell.RowIndex;
                        int firstCol = selectedCell.ColumnIndex;

                        for (int i = 0; i < cellDataSet.SelectedCells.Count; i++)
                        {
                            DataCell cellProps = cellDataSet.SelectedCells[i] as DataCell;


                            SheetCell targetCell = CurrentLayout.Rows[selectedRow].Cells[selectedColumn] as SheetCell;
                            
                            

                            if(cellProps.CellRow != lastRow)
                            {
                                targetCell.Value = cellProps.CellValue;
                                targetCell.BackGroundColor = cellProps.BackGroundColor;
                                targetCell.ForeColor = cellProps.ForeColor;
                                targetCell.Alignment = cellProps.Alignment;
                                selectedRow++;
                                selectedColumn = firstCol;
                                lastRow++;
                            }
                            else
                            {
                                targetCell.Value = cellProps.CellValue;
                                targetCell.BackGroundColor = cellProps.BackGroundColor;
                                targetCell.ForeColor = cellProps.ForeColor;
                                targetCell.Alignment = cellProps.Alignment;
                                selectedColumn++;
                            }



                            CurrentLayout.InvalidateCell(targetCell);
                         
                        }
         


                    }
                    else
                    {
                        MessageBox.Show("Retrieved object is null. Check serialization1.");
                    }
                    
                }
            }

               
        }



        private void copyBtn_Click(object sender, EventArgs e)
        {
            if (CurrentLayout.SelectedCells.Count is int numberOfSelectedCells && numberOfSelectedCells > 0)
            {
                if (numberOfSelectedCells > 1)
                {
                    DataCellSet dataCellSet = new DataCellSet();

                    var selectedGridCells = CurrentLayout.SelectedCells;
               

                    for(int i = 0; i < selectedGridCells.Count; i++)
                    {
                        SheetCell selectedCell = selectedGridCells[i] as SheetCell;
                        int cellRow = selectedCell.RowIndex;
                        int cellCol = selectedCell.ColumnIndex;
                        string cellValue = selectedCell.EditedFormattedValue.ToString();
                        System.Drawing.Color backGroundColor = selectedCell.BackGroundColor;
                        System.Drawing.Color foreColor = selectedCell.ForeColor;
                        System.Drawing.StringAlignment alignment = selectedCell.Alignment;

                        SheetCell.DataCell dataCell = new SheetCell.DataCell(cellRow, cellCol, cellValue, backGroundColor, foreColor, alignment);

                        dataCellSet.SelectedCells.Add(dataCell);
                        
                    }

                    System.Windows.Forms.DataObject dataObject = new System.Windows.Forms.DataObject("SheetCellDataSet", dataCellSet);

                    if (dataObject != null)
                    {
                        Clipboard.Clear();
                        Clipboard.SetDataObject(dataObject);
                    }
                }
                else
                {
                    //Only one cell selected
                    SheetCell selectedCell = CurrentLayout.SelectedCells[0] as SheetCell;
                    int cellRow = selectedCell.RowIndex;
                    int cellCol = selectedCell.ColumnIndex;
                    string cellValue = selectedCell.EditedFormattedValue.ToString();
                    System.Drawing.Color backGroundColor = selectedCell.BackGroundColor;
                    System.Drawing.Color foreColor = selectedCell.ForeColor;
                    System.Drawing.StringAlignment alignment = selectedCell.Alignment;

                    SheetCell.DataCell dataCell = new SheetCell.DataCell(cellRow, cellCol,cellValue, backGroundColor, foreColor, alignment);

                    System.Windows.Forms.DataObject dataObject = new System.Windows.Forms.DataObject("SheetCellData", dataCell);
                    if (dataObject != null)
                    {
                        Clipboard.Clear();
                        Clipboard.SetDataObject(dataObject);
                    }
                }
            }
        }

        private void cutBtn_Click(object sender, EventArgs e)
        {
            copyBtn_Click(sender, e);
            SheetCell selectedCell = CurrentLayout.SelectedCells[0] as SheetCell;
            selectedCell.BackGroundColor = System.Drawing.Color.White;
            selectedCell.ForeColor = System.Drawing.SystemColors.InfoText;
            selectedCell.Alignment = System.Drawing.StringAlignment.Far;
            
            CurrentLayout.InvalidateCell(selectedCell);

        }
    }
}
