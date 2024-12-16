using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Net.Mail;


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

        private SpreadsheetApp initialForm; // Patricia - to be used to close the current form, not the full aplication
        public Document(string name, int numOfRows, int numOfColumns, string filePath, SpreadsheetApp initialForm) // SpreadsheetApp initialForm - Patricia
        {
            FileName = name;
            NumOfRows = numOfRows;
            NumOfColumns = numOfColumns;
            FilePath = filePath;
            OriginDate = DateTime.Now;
            LastModificationDate = DateTime.Now;
            CurrentDataTable = new DataSource(numOfRows, numOfColumns);
            CurrentLayout = new Sheet(CurrentDataTable);

            /* Begin - Patricia */
            CurrentLayout.KeyDown += CurrentLayout_KeyDown; // To manipulate windows keys freely
            CurrentLayout.MultiSelect = true; // permit multi selection of cell 
            CurrentLayout.SelectionMode = DataGridViewSelectionMode.CellSelect;  // permit multi selection of cell
            CurrentLayout.SelectionChanged += CurrentLayout_SelectionChanged; // to permit moving arrows for all sides so that user can select more than 1 cell by using <shift> + arrows 
            /* End - Patricia */

            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();

            this.initialForm = initialForm; // Patricia

        }
        public Document(string name, int numOfRows, int numOfColumns, string filePath, SpreadsheetDocument fileToBeOpened, SpreadsheetApp initialForm)
        {
            FileName = name;
            NumOfRows = numOfRows;
            NumOfColumns = numOfColumns;
            FilePath = filePath;
            OriginDate = DateTime.Now;
            LastModificationDate = DateTime.Now;
            CurrentDataTable = new DataSource(fileToBeOpened, numOfRows, numOfColumns);

            /* Begin - Patricia */
            CurrentLayout.KeyDown += CurrentLayout_KeyDown; // to permit deleting values typed wrongly
            CurrentLayout.MultiSelect = true; // permit multi selection of cell 
            CurrentLayout.SelectionMode = DataGridViewSelectionMode.CellSelect;  // permit multi selection of cell
            CurrentLayout.SelectionChanged += CurrentLayout_SelectionChanged; // to permit moving arrows for all sides so that user can select more than 1 cell by using <shift> + arrows 
            /* End - Patricia */

            DataTables = new List<DataTable>() { CurrentDataTable };
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();

            this.initialForm = initialForm; // Patricia
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
        
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e) // Giovanni & Patricia
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
             //   sfd.FileName = "Sheet.xlsx";
                sfd.Title = "Save File As";
                sfd.FileName = this.FileName;
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|PDF Files (*.pdf)|*.pdf";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string extension = Path.GetExtension(sfd.FileName).ToLower(); // Get file extension
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
        private void saveToolStripMenuItem_Click(object sender, EventArgs e) // Giovanni
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
            Application.Exit();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e) // Patricia - close section and back to main screen
        {
            this.Close();

            initialForm.Show();
        }

        private void CurrentLayout_KeyDown(object sender, KeyEventArgs e) // Patricia - Manipulation of keys like Op.System Windows does
        {
            // Allow DataGridView works multi selection of cells by using <SHIFT> + Arrows
            if (e.Shift && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                e.Handled = false; 
                return;
            }

            // Allow user to move freely over cells 
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
            {
                e.Handled = false;
                return;
            }

            // Deselect all cells when <ESC> is pressed
            if (e.KeyCode == Keys.Escape)
            {
                CurrentLayout.ClearSelection(); // Remove selection of all cells
                e.Handled = true;
                return;
            }

            // Clean cells and delete by using <BACKSPACE> or <DEL>
            foreach (DataGridViewCell cell in CurrentLayout.SelectedCells)
            {
                if (!cell.ReadOnly)
                {
                    if (e.KeyCode == Keys.Delete) // clean cell completely
                    {
                        cell.Value = "";
                    }
                    else if (e.KeyCode == Keys.Back) // Backspace - delete one char by time
                    {
                        if (cell.Value != null && cell.Value is string cellValue && cellValue.Length > 0)
                        {
                            // delete the last char
                            cell.Value = cellValue.Substring(0, cellValue.Length - 1);
                        }
                    }
                }
            }

            e.Handled = true;
        }

        private void CurrentLayout_SelectionChanged(object sender, EventArgs e) // Patricia
        {
            try
            {
                // Get selected cells from DataGridViewCell to calculate the average
                var selectedCells = CurrentLayout.SelectedCells.Cast<DataGridViewCell>();

                // Filter only cells with valid numeric values
                var numericValues = selectedCells
                    .Where(cell => double.TryParse(cell.Value?.ToString(), out _))
                    .Select(cell => double.Parse(cell.Value.ToString()));

                if (!numericValues.Any())
                {
                    Console.WriteLine("No numeric cell was selected");
                    return;
                }

                double average = numericValues.Average();

                Console.WriteLine($"The average is: {average:F2}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error by calculating average: {ex.Message}");
            }
        }
    }
}
