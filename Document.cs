using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;


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
            //CurrentDataTable = CurrentDataTable.TransferDataToTable(fileToBeOpened, numOfRows, numOfColumns);
            CurrentLayout = new Sheet(CurrentDataTable);

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

        private DataTable TransferDataToTable(SpreadsheetDocument openedFile)
        {
            DataTable tableToFill = new DataTable();
            tableToFill = AddColumnHeaderToTable(tableToFill);

            WorkbookPart workbookPart = openedFile.WorkbookPart;
            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault();

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            var rows = worksheetPart.Worksheet.Descendants<Row>();

            foreach (Row row in rows)
            {
                DataRow dataRow = tableToFill.NewRow();
                int columnIndex = 0;

                foreach (Cell cell in row.Descendants<Cell>())
                {
                    string cellValue = GetCellValue(workbookPart, cell);
                    if (columnIndex < tableToFill.Columns.Count)
                    {
                        dataRow[columnIndex] = cellValue;
                    }
                    columnIndex++;
                }
                tableToFill.Rows.Add(dataRow);
            }
            return tableToFill;
        }


        private DataTable AddColumnHeaderToTable(DataTable table)   // uncommented by Patricia to test
        {
            DataColumn Column;
            string columnName = "";
            for (int i = 1; i < NumOfColumns; i++)
            {
                if (i <= 26)
                {
                    // First 26 columns are just A-Z.
                    columnName = $"{(char)(i + 64)}"; // 'A' is 65 in ASCII, so adding 64 to get A-Z.
                    Column = new DataColumn(columnName);
                }
                else
                {
                    // For columns beyond Z (i.e., AA, AB, etc.)
                    int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
                    int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
                    // Combine the prefix and suffix to get AA, AB, etc.
                    columnName = $"{(char)(quotient + 64)}{(char)(remainder + 64)}";
                    Column = new DataColumn(columnName);
                }
                Column.DataType = typeof(string);
                Column.AllowDBNull = true;
                Column.DefaultValue = "";
                Column.MaxLength = 255;
                table.Columns.Add(Column);
            }
            return table;
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
        //private void saveAsToolStripMenuItem_Click(object sender, EventArgs e) // comented by Patricia
        //{

        //    using (SaveFileDialog sfd = new SaveFileDialog())
        //    {
        //        //sfd.FileName = "";
        //        //sfd.Filter = "Excel|*.xlsx";
        //        sfd.FileName = "Sheet.xlsx"; // Patricia - teste
        //                                     //   sfd.Filter = "Excel Files (*.xlsx)|*.xlsx"; // Patricia - teste
        //        sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|PDF Files (*.pdf)|*.pdf"; // Patricia teste

        //        if (sfd.ShowDialog() == DialogResult.OK)
        //        {
        //            // USING OPENXML
        //            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(sfd.FileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        //            {
        //                WorkbookPart workbookPart = doc.AddWorkbookPart();
        //                workbookPart.Workbook = new Workbook();
        //                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //                worksheetPart.Worksheet = new Worksheet(new SheetData());
        //                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
        //                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };


        //                sheets.Append(sheet);
        //                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        //                // Write DataTable rows
        //                foreach (DataRow dataRow in this.CurrentDataTable.Rows)
        //                {
        //                    Row newRow = new Row();
        //                    foreach (var item in dataRow.ItemArray)
        //                    {
        //                        Cell cell = new Cell
        //                        {
        //                            DataType = CellValues.String,
        //                            CellValue = new CellValue(item.ToString())
        //                        };
        //                        newRow.AppendChild(cell);
        //                    }
        //                    sheetData.AppendChild(newRow);
        //                }
        //                workbookPart.Workbook.Save();
        //            }
        //        }
        //    }
        //}

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e) // Patricia - SAVE AS
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.FileName = "Sheet.xlsx";
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|PDF Files (*.pdf)|*.pdf";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string extension = Path.GetExtension(sfd.FileName).ToLower(); // Get file extension

                    if (extension == ".xlsx")
                    {
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
                                        CellValue = new CellValue(item?.ToString() ?? string.Empty) // Make sure all cells be processed
                                    };
                                    newRow.AppendChild(cell);
                                }
                                sheetData.AppendChild(newRow);
                            }
                        }
                    }
                    else if (extension == ".csv")
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName))
                        {
                            var columnNames = CurrentDataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                            sw.WriteLine(string.Join(",", columnNames));

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

        private void pasteBtn_Click(object sender, EventArgs e)
        {

        }

        private void copyBtn_Click(object sender, EventArgs e)
        {

        }

        private void cutBtn_Click(object sender, EventArgs e)
        {

        }

        private void CurrentLayout_SelectionChanged(object sender, EventArgs e)
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
