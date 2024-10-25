using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    internal class Sheet : DataGridView
    {

        public Sheet(DataTable source)
        {
            ((System.ComponentModel.ISupportInitialize)this).BeginInit();

            this.Name = source.Namespace;
            this.TabIndex = 14;
            this.Dock = DockStyle.Fill;
            this.BackgroundColor = Color.White;

            //ALLOW USER
            this.AllowUserToDeleteRows = true;
            this.AllowUserToAddRows = true;
            this.AllowUserToResizeColumns = true;

            //HEADERS
            this.EnableHeadersVisualStyles = false; // true prevents custom style on headers

            //ROW HEADERS
            this.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.RowHeadersDefaultCellStyle.BackColor = Color.Gray;
            this.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            this.RowHeadersWidth = 60;
            this.RowPostPaint += (object sender, DataGridViewRowPostPaintEventArgs e) => {
                using (Brush sb = new SolidBrush(Color.Black))
                {
                    e.Graphics.DrawString($"{(e.RowIndex)}", new Font("Verdant", 10), sb, new PointF(e.RowBounds.X + 15, e.RowBounds.Y + 5));
                }
            };
            //COLUMN HEADERS
            this.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ColumnHeadersDefaultCellStyle.Font = new Font(FontFamily.GenericSansSerif, 8, FontStyle.Regular);
            this.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            this.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            this.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //CELLS
            this.DefaultCellStyle.SelectionBackColor = Color.Wheat;
            this.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            //Grid Generation
            foreach (DataColumn column in source.Columns)
            {
                this.Columns.Add(new SheetColumn()
                {
                    HeaderText = column.ColumnName,
                });
            }
            for (int row = 0; row < source.Rows.Count; row++)
            {
                this.Rows.Add(new SheetRow()); //adding custom row.

                // Populate cells in the row
                for (int col = 0; col < source.Columns.Count; col++)
                {
                    // Use custom thisCell
                    SheetCell cell = new SheetCell();
                    cell.Value = source.Rows[row][col].ToString();

                    // Assign custom cell to the DataGridView row
                    this.Rows[row].Cells[col] = cell;

                }
            }

            ((System.ComponentModel.ISupportInitialize)this).EndInit();

            this.CellEndEdit += (object sender, DataGridViewCellEventArgs e) =>
            {
                DataGridView grid = sender as DataGridView;
                SheetCell editedCell = (SheetCell)grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                editedCell.SetValue(grid, e);
            };
        }
    }
}
