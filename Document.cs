using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static spreadsheetApp.SheetCell;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Path = System.IO.Path;


namespace spreadsheetApp
{
    public partial class Document : Form
    {
        public string FileName { get; set; } = " ";
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
                sfd.FileName = this.FileName;
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|PDF Files (*.pdf)|*.pdf";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string extension = System.IO.Path.GetExtension(sfd.FileName).ToLower(); // Get file extension
                    FilePath = sfd.FileName;
                    if (extension == ".xlsx")
                    {
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(sfd.FileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)) // USING OPENXML
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
                    else if (extension == ".csv")
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName))
                        {
                            //  var columnNames = CurrentDataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                            //  sw.WriteLine(string.Join(",", columnNames));

                            // Write data of rows
                            foreach (DataRow row in CurrentDataTable.Rows)
                            {
                                var fields = row.ItemArray.Select(field => field?.ToString()?.Replace(",", "\"\"") ?? string.Empty); // Check commas and null values
                                sw.WriteLine(string.Join(",", fields));
                            }
                        }
                    }
                    else if (extension == ".pdf")
                    {
                        using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
                            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, fs);

                            pdfDoc.Open();
                            PdfPTable table = new PdfPTable(CurrentDataTable.Columns.Count)
                            {
                                WidthPercentage = 100, // Adjust table to page size
                                SpacingBefore = 10f,
                                SpacingAfter = 10f
                            };

                            // Add header of columns
                            foreach (DataColumn column in CurrentDataTable.Columns)
                            {
                                PdfPCell headerCell = new PdfPCell(new iTextSharp.text.Phrase(column.ColumnName))
                                {
                                    HorizontalAlignment = Element.ALIGN_CENTER,
                                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY
                                };
                                table.AddCell(headerCell);
                            }

                            // Add data of cells
                            foreach (DataRow row in CurrentDataTable.Rows)
                            {
                                for (int colIndex = 0; colIndex < CurrentDataTable.Columns.Count; colIndex++)
                                {
                                    // Verifiy cell value or substitute for empty value
                                    string cellValue = row[colIndex]?.ToString() ?? " ";
                                    table.AddCell(new iTextSharp.text.Phrase(cellValue));
                                }
                            }

                            pdfDoc.Add(table);
                            pdfDoc.Close();
                        }

                    }
                    else
                    {
                        MessageBox.Show("File format not supported", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                saveAsToolStripMenuItem_Click(sender, e);
            }
            else
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(FilePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)) // USING OPENXML
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
                if (string.IsNullOrEmpty(title)) throw new Exception("Invalid document title");
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

        private void avgMenuItem_Click(object sender, EventArgs e) // Patricia
        {
            try
            {
                var selectedCells = CurrentLayout.SelectedCells.Cast<DataGridViewCell>();

                var numericValues = selectedCells
                    .Where(cell => double.TryParse(cell.Value?.ToString(), out _))
                    .Select(cell => double.Parse(cell.Value.ToString()));

                if (!numericValues.Any())
                {
                    MessageBox.Show("No numeric value was selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Calculate average
                double average = numericValues.Average();

                MessageBox.Show($"Average: {average}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in calculating the average: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void sumMenuItem_Click(object sender, EventArgs e) // Patricia
        {
            try
            {
                var selectedCells = CurrentLayout.SelectedCells.Cast<DataGridViewCell>();

                var numericValues = selectedCells
                    .Where(cell => double.TryParse(cell.Value?.ToString(), out _))
                    .Select(cell => double.Parse(cell.Value.ToString()));

                if (!numericValues.Any())
                {
                    MessageBox.Show("No numeric value was selected", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Calculate summation
                double sum = numericValues.Sum();

                MessageBox.Show($"Summation: {sum}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in calculating the summation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) // Patricia - close application
        {
            this.Close();
            Application.Exit();
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
                else if (clipboardDataObject.GetDataPresent("SheetCellDataSet"))
                {

                    SheetCell.DataCellSet cellDataSet = clipboardDataObject.GetData("SheetCellDataSet") as SheetCell.DataCellSet;
                    if (cellDataSet != null)
                    {
                        List<SheetCell> cellsToBePasted = new List<SheetCell>();

                        SheetCell selectedCell = CurrentLayout.SelectedCells[0] as SheetCell;

                        int selectedRow = selectedCell.RowIndex;
                        int selectedColumn = selectedCell.ColumnIndex;

                        int lastRow = selectedCell.RowIndex;
                        int firstCol = selectedCell.ColumnIndex;

                        for (int i = 0; i < cellDataSet.SelectedCells.Count; i++)
                        {
                            DataCell cellProps = cellDataSet.SelectedCells[i] as DataCell;


                            SheetCell targetCell = CurrentLayout.Rows[selectedRow].Cells[selectedColumn] as SheetCell;



                            if (cellProps.CellRow != lastRow)
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


                    for (int i = 0; i < selectedGridCells.Count; i++)
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

                    SheetCell.DataCell dataCell = new SheetCell.DataCell(cellRow, cellCol, cellValue, backGroundColor, foreColor, alignment);

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
            selectedCell.Value = String.Empty;
            CurrentLayout.InvalidateCell(selectedCell);

        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            closeToolStripMenuItem_Click((object)sender, e);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Open File";
                ofd.FileName = "";
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                //filePath = ofd.FileName;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // USING OPENXML
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ofd.FileName, false))
                    {
                        WorkbookPart workbookPart = doc.WorkbookPart;
                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet1 = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First();
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet1.Id);
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                        NumOfRows = worksheetPart.Worksheet.Descendants<Row>().Count();
                        Row FirstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();
                        if (FirstRow != null)
                        {
                            NumOfColumns = FirstRow.Descendants<Cell>().Count();
                        }
                        else
                        {
                            NumOfColumns = 0;
                        }
                        if (NumOfColumns == 0 && NumOfRows == 0)
                        {
                            Document newOpenedDocument = new Document(ofd.FileName, 30, 10, FilePath);
                            newOpenedDocument.Display();
                        }
                        else
                        {
                            Document newOpenedDocument = new Document(ofd.FileName, NumOfRows, NumOfColumns, FilePath, doc);
                            newOpenedDocument.Display();
                        }
                    }
                }
            }
        }
    }
}
