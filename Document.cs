using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        public string FilePath { get;  set; }
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

            CurrentDataTable = CreateEmptyTable();
            DataTables = new List<DataTable>() { CurrentDataTable };
            CurrentLayout = CreateLayoutFrom(CurrentDataTable);
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

            CurrentDataTable = TransferDataToTable(fileToBeOpened);
            DataTables = new List<DataTable>() { CurrentDataTable };
            CurrentLayout = CreateLayoutFrom(CurrentDataTable);
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }
        public Document(string name, int cols, int rows)
        {
            CurrentDataTable = CreateEmptyTable();
            DataTables = new List<DataTable>() { CurrentDataTable };
            CurrentLayout = CreateLayoutFrom(CurrentDataTable);
            Layouts = new List<DataGridView>() { CurrentLayout };
            DisplayLayout(CurrentLayout);
            InitializeComponent();
        }
        // ---------------------------------------------- LAYOUTLOGIC GUI LOGIC DATAGRIDVIEW ----------------------------------------
        private DataGridView CreateLayoutFrom(DataTable table)
        {
            DataGridView Sheet = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)Sheet).BeginInit();
            
            Sheet.Name = "sheet1";           
            Sheet.TabIndex = 14;
            Sheet.Dock = DockStyle.Fill;
            Sheet.BackgroundColor = System.Drawing.Color.White;
            Sheet.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            
            Sheet.EnableHeadersVisualStyles = false;
            Sheet.RowHeadersWidth = 60;
            
            Sheet.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            Sheet.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;           
            Sheet.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(System.Drawing.FontFamily.GenericSansSerif, 8, FontStyle.Regular);
            Sheet.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            Sheet.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            Sheet.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Sheet.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            Sheet.AllowUserToDeleteRows = true;
            Sheet.AllowUserToAddRows = true;
            Sheet.AllowUserToResizeColumns = true;
            
            ((System.ComponentModel.ISupportInitialize)Sheet).EndInit();
            Sheet.DataSource = table;

            return Sheet;
        }

        // ---------------------------------------------- DATA LOGIC BUSINESS LOGIC DATATABLE VIRTUAL SHEET ----------------------------------------
        private DataTable CreateEmptyTable() // just emptied the arguments, we can access the props straightforward.
        {
            DataTable table = new DataTable("sheet 1");
            table = AddColumnHeaderToTable(table);
            return table;
        }
        private DataTable AddColumnHeaderToTable(DataTable table)
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
        private DataTable TransferDataToTable(SpreadsheetDocument openedFile)
        {
            DataTable tableToFill = new DataTable();
            tableToFill = AddColumnHeaderToTable(tableToFill);

            WorkbookPart workbookPart = openedFile.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
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
            this.ShowDialog();
        }

        // adding columns and rows according to user input.
        //public void CreateGrid ()
        //{
        //    DataGridView dataGrid = new DataGridView(); // creating a datagrid.
        //    dataGrid.Size = new Size(1000, 300); // setting its size.
        //    dataGrid.Location = new Point(0, 60); // setting its location.

        //    DataTable dataTable = new DataTable(); // creating a datatable will make our life easier later on.

        //    // adding columns to our table. First column act as rows header.
        //    dataTable.Columns.Add("*"); 
        //    dataTable.Columns.Add("A");
        //    dataTable.Columns.Add("B");
        //    dataTable.Columns.Add("C");
        //    dataTable.Columns.Add("D");

        //    // adding rows to the DataTable.
        //    dataTable.Rows.Add("1");
        //    dataTable.Rows.Add("2");
        //    dataTable.Rows.Add("3");
        //    dataTable.Rows.Add("4");
        //    dataTable.Rows.Add("5");
        //    dataTable.Rows.Add("6");
        //    dataTable.Rows.Add("7");
        //    dataTable.Rows.Add("8");

        //    dataGrid.DataSource = dataTable; // attaching datatable to the datagrid.
        //    this.Controls.Add(dataGrid); // adding grid to list of controls.

        //    dataGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // centering the values.
        //    dataGrid.RowHeadersVisible = false; // removing first column so the next one will act as row headers.
        //    //dataGrid.Columns[].Width = 20;
        //    //dataGrid.AllowUserToAddRows = false; // a new row is added by default.

        //    dataGrid.Columns[1].Frozen = true; // freezing the first column is not working, further investigation.
        //    //dataGrid.Rows[0].HeaderCell.Value = ""; // also not working, don't know why
        //    dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // Auto-size columns, could not do the same for rows.
        //}
    }
}
