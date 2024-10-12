using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        public string name {  get; set; }
        public int numOfRows { get; set; }
        public int numOfColumns { get; set; }
        //public Document(string name, int numOfRows, int numOfColumns)
        public Document()
        {
            InitializeComponent();
            //CreateGrid(numOfRows, numOfColumns);
            CreateGrid();
        }
        public void Display()
        {
            this.Show();
        }

        // adding columns and rows according to user input.
        public void CreateGrid ()
        {
            DataGridView dataGrid = new DataGridView(); // creating a datagrid.
            dataGrid.Size = new Size(1000, 300); // setting its size.
            dataGrid.Location = new Point(0, 60); // setting its location.

            DataTable dataTable = new DataTable(); // creating a datatable will make our life easier later on.

            // adding columns to our table. First column act as rows header.
            dataTable.Columns.Add("*"); 
            dataTable.Columns.Add("A");
            dataTable.Columns.Add("B");
            dataTable.Columns.Add("C");
            dataTable.Columns.Add("D");

            // adding rows to the DataTable.
            dataTable.Rows.Add("1");
            dataTable.Rows.Add("2");
            dataTable.Rows.Add("3");
            dataTable.Rows.Add("4");
            dataTable.Rows.Add("5");
            dataTable.Rows.Add("6");
            dataTable.Rows.Add("7");
            dataTable.Rows.Add("8");

            dataGrid.DataSource = dataTable; // attaching datatable to the datagrid.
            this.Controls.Add(dataGrid); // adding grid to list of controls.

            dataGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // centering the values.
            dataGrid.RowHeadersVisible = false; // removing first column so the next one will act as row headers.
            //dataGrid.Columns[].Width = 20;
            //dataGrid.AllowUserToAddRows = false; // a new row is added by default.

            dataGrid.Columns[1].Frozen = true; // freezing the first column is not working, further investigation.
            //dataGrid.Rows[0].HeaderCell.Value = ""; // also not working, don't know why
            dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // Auto-size columns, could not do the same for rows.
        }
    }
}
