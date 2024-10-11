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

        public string Name { get; set; }
        public DateTime OriginDate { get; set; }
        public DateTime LastModificationDate { get; set; }

        public List<DataTable> DataTables { get; set; }
        public List<DataGridView> Layouts { get; set; }
        public DataGridView CurrentLayout { get; set; }

        public DataTable CurrentDataTable { get; set; }


        public Document()
        {
            
            CurrentDataTable = CreateEmptyTable();
            DataTables = new List<DataTable>() { CurrentDataTable };
            CurrentLayout = CreateLayoutFrom(CurrentDataTable);
            Layouts = new List<DataGridView>() { CurrentLayout };
            InitializeComponent();
        }
        private DataGridView CreateLayoutFrom(DataTable table)
        {
            DataGridView Sheet = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)Sheet).BeginInit();
            Sheet.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            Sheet.Name = "sheet1";
            Sheet.RowHeadersWidth = 62;
            Sheet.TabIndex = 14;
            Sheet.Dock = DockStyle.Fill;
            ((System.ComponentModel.ISupportInitialize)Sheet).EndInit();
            Sheet.DataSource = table;
            Controls.Add(Sheet);



            return Sheet;
        }

        private DataTable CreateEmptyTable(int columns=200, int rows=200)
        {
            DataTable Table = new DataTable("sheet 1");
            DataColumn _column;
            string columnName = "";

            for (int i = 1; i < columns; i++)
            {
                if (i <= 26)
                {
                    // First 26 columns are just A-Z

                    columnName = $"{(char)(i + 64)}"; // 'A' is 65 in ASCII, so adding 64 to get A-Z
                    _column = new DataColumn(columnName);
                }
                else
                {
                    // For columns beyond Z (i.e., AA, AB, etc.)
                    int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
                    int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)

                    // Combine the prefix and suffix to get AA, AB, etc.
                    columnName = $"{(char)(quotient + 64)}{(char)(remainder + 64)}";
                    _column = new DataColumn(columnName);
                }


                _column.DataType = typeof(string);
                _column.AllowDBNull = true;
                _column.DefaultValue = "";
                _column.MaxLength = 255;

            

                Table.Columns.Add(_column);
            }
            return Table;
        }


        private void DisplayLayout()
        {

        }

        public void Display()
        {
            this.Show();
        }

 
    }
}
