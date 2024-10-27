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

            Name = source.Namespace;
            TabIndex = 14;
            Dock = DockStyle.Fill;
            BackgroundColor = Color.White;

            //ALLOW USER
            AllowUserToDeleteRows = true;
            AllowUserToAddRows = true;
            AllowUserToResizeColumns = true;

            // this shuts downs the default style.
            EnableHeadersVisualStyles = false;

            //COLUMN HEADERS - Default HeaderClass from Grid... needs inheritances for customization

            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            ColumnHeadersDefaultCellStyle.Font = new Font(FontFamily.GenericSansSerif, 8, FontStyle.Regular);
            ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            //ROW HEADERS
            RowHeadersWidth = 50;
            RowHeadersDefaultCellStyle.BackColor = Color.LightGray;
            RowHeadersDefaultCellStyle.SelectionBackColor = Color.LightGray;

            //Grid Generation
            foreach (DataColumn column in source.Columns) Columns.Add(new SheetColumn() { HeaderText = column.ColumnName, });

            foreach (DataRow row in source.Rows) Rows.Add();



            ((System.ComponentModel.ISupportInitialize)this).EndInit();

            // EVENTS

            RowPostPaint += (sender, e) =>
            {
                using (Brush sb = new SolidBrush(Color.Black))
                {
                    int xOffSet = 15;
                    if (e.RowIndex > 9) xOffSet = 10;
                    e.Graphics.DrawString($"{e.RowIndex}", new Font("Verdant", 10), sb, new PointF(e.RowBounds.X + xOffSet, e.RowBounds.Y + 5));
                }
            };

            ColumnHeaderMouseClick += (sender, e) =>
            {
                DataGridView grid = sender as DataGridView;

                // Check that a valid column index is clicked
                if (e.ColumnIndex >= 0)
                {
                    grid.ClearSelection(); // Clear any existing selection
                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        row.Cells[e.ColumnIndex].Selected = true; // Select each cell in the clicked column
                    }
                }
            };

            CellValidating += (sender, cell) =>
            {
                DataGridView grid = sender as DataGridView;
                SheetCell editedCell = (SheetCell)grid.Rows[cell.RowIndex].Cells[cell.ColumnIndex];
                editedCell.SetValue(grid, cell);

                //Updating source
                source.Rows[editedCell.RowIndex][editedCell.ColumnIndex] = editedCell.Value;

                //How to update the cell format so it can be saved n loaded with the cells values???

            };
        }
    }
}
