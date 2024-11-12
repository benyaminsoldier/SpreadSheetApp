using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            AllowUserToDeleteRows = false;
            AllowUserToAddRows = false;
            AllowUserToResizeColumns = true;

            // this shuts downs the default style.
            EnableHeadersVisualStyles = false;
            CellBorderStyle = DataGridViewCellBorderStyle.Raised;   

            //COLUMN HEADERS - Default HeaderClass from Grid... needs inheritances for customization

            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            ColumnHeadersDefaultCellStyle.Font = new Font(FontFamily.GenericSansSerif, 8, FontStyle.Regular);
            ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;

            //ROW HEADERS
            RowTemplate = new SheetRow();
            RowHeadersWidth = 50;
            RowHeadersDefaultCellStyle.BackColor = Color.LightGray;
            RowHeadersDefaultCellStyle.SelectionBackColor = Color.LightGray;

            DefaultCellStyle.SelectionBackColor = Color.Transparent;
            DefaultCellStyle.SelectionForeColor = Color.Red;
            
            
            //Grid Generation
            //Try block missed here to handle unmeasured user table sizes
            
            foreach (DataColumn column in source.Columns) Columns.Add(new SheetColumn() { HeaderText = column.ColumnName, });

            foreach (DataRow row in source.Rows) Rows.Add(new SheetRow());

            // EVENTS

            
            EditingControlShowing += ( sender, e) =>
            {
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            };

            CellValidating += (sender, cell) =>
            {
                Sheet grid = sender as Sheet;
                SheetCell editedCell = (SheetCell)grid.Rows[cell.RowIndex].Cells[cell.ColumnIndex];
               
                editedCell.SetValue(grid, cell);
                this.RefreshEdit(); 
                this.Refresh();

                //Updating source
                source.Rows[editedCell.RowIndex][editedCell.ColumnIndex] = editedCell.Value;

                //How to update the cell format so it can be saved n loaded with the cells values???

            };

            CellEnter += (s, e) => { this.InvalidateCell(this.CurrentCell); this.Refresh(); };

            CellPainting += (sender, e) =>
            {
                //System.Drawing.Drawing2D.SmoothingMode sm = SmoothingMode.AntiAlias;
                
                SheetCell cellToBePainted;
                bool cellIsWithinBounds = e.RowIndex >= 0 && e.RowIndex < this.Rows.Count && e.ColumnIndex >= 0 && e.ColumnIndex < this.Columns.Count;

                if (cellIsWithinBounds)
                {
                    
                    bool selectedCell = this.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected;

                    if (selectedCell)
                    {
                        cellToBePainted = this.Rows[e.RowIndex].Cells[e.ColumnIndex] as SheetCell;
                        var cellToBePainted2 = cellToBePainted as DataGridViewTextBoxCell;

                        if (cellToBePainted.IsFormatted)
                        {

                            cellToBePainted.DrawSelectedCellBorder(e);
                            cellToBePainted.FillCellBackGround2(cellToBePainted.BackGroundColor);
                            cellToBePainted.DrawCellString(e);
                            


                        }
                        else
                        {
                            cellToBePainted.DrawSelectedCell(e);
                            
                        }
                        
                    }
                    else
                    {
                        cellToBePainted = this.Rows[e.RowIndex].Cells[e.ColumnIndex] as SheetCell;

                        if (cellToBePainted.IsFormatted)
                        {
                            cellToBePainted.FillCellBackGround(cellToBePainted.BackGroundColor);
                            cellToBePainted.DrawDefaultCellBorder(e);
                            cellToBePainted.DrawCellString(e);
                            
                            //cellToBePainted.DrawSelectedBorders(e);
                            
                        }
                        else
                        {
                            cellToBePainted.DrawNonSelectedCell(e);
                            
                        }

                    }
                    e.Handled = true;
                }
                else e.Handled = false;

                
            };
          

            RowPostPaint += (sender, e) =>
            {
                using (Brush sb = new SolidBrush(Color.Black))
                {                   
                    int xOffSet = 15;
                    if (e.RowIndex > 9) xOffSet = 10;
                    e.Graphics.DrawString($"{e.RowIndex+1}", new Font("Verdant", 10), sb, new PointF(e.RowBounds.X + xOffSet, e.RowBounds.Y + 5));
                }
            };

            ColumnHeaderMouseClick += (sender, e) =>
            {
                Sheet grid = sender as Sheet;

                // Check that a valid column index is clicked
                if (e.ColumnIndex >= 0)
                {
                    grid.ClearSelection(); // Clear any existing selection
                    foreach (SheetRow row in grid.Rows)
                    {
                        row.Cells[e.ColumnIndex].Selected = true; // Select each cell in the clicked column
                    }
                }
            };



            ((System.ComponentModel.ISupportInitialize)this).EndInit();

        }
    }
}
