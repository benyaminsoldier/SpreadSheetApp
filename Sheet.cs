using System;
using System.Collections.Generic;
using System.Data;
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
                DataGridView grid = sender as DataGridView;
                SheetCell editedCell = (SheetCell)grid.Rows[cell.RowIndex].Cells[cell.ColumnIndex];
               
                editedCell.SetValue(grid, cell);
                this.RefreshEdit(); 
                this.Refresh();

                //Updating source
                source.Rows[editedCell.RowIndex][editedCell.ColumnIndex] = editedCell.Value;

                //How to update the cell format so it can be saved n loaded with the cells values???

            };

            //CellEnter += (s, e) => { this.InvalidateCell(this.CurrentCell); };

            CellPainting += ( sender, e) =>
            {
                //e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                // For all regular cells within grid
                bool cellWithinBounds = e.RowIndex >= 0 && e.RowIndex < this.Rows.Count && e.ColumnIndex >= 0 && e.ColumnIndex < this.Columns.Count;
                
                if (cellWithinBounds)
                {
                    bool selectedCell = this.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected;
                    using (SolidBrush stringBrush = new SolidBrush(this.ForeColor))
                    {
                        using (SolidBrush backBrush = new SolidBrush(Color.White))
                        {
                            if (selectedCell)
                            {


                                using (Pen borderpen = new Pen(Color.DarkOliveGreen, 2))
                                {
                                    Rectangle rectDimensions = e.CellBounds;
                                    rectDimensions.Width -= 1;
                                    rectDimensions.Height -= 1;
                                    rectDimensions.X = rectDimensions.X + 1;
                                    rectDimensions.Y = rectDimensions.Y + 1;


                                    e.Graphics.FillRectangle(backBrush, rectDimensions);

                                    e.Graphics.DrawString(
                                        e.FormattedValue as string,
                                        this.Font,
                                        stringBrush,
                                        rectDimensions,
                                        new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far }
                                    );


                                    if (!String.IsNullOrEmpty(this.CurrentCell.ErrorText))
                                    {
                                        TextBox txtBox = this.EditingControl as TextBox;
                                        this.BeginEdit(true);
                                        txtBox.SelectionStart = int.Parse(this.CurrentCell.ErrorText);
                                        txtBox.SelectionLength = 1;
                                        txtBox.Focus();
                                    }

                                    e.Graphics.DrawRectangle(borderpen, rectDimensions);
                                    e.Handled = true;

                                }
                            }
                            else 
                            {

                                using (Pen eraser = new Pen(this.BackgroundColor, 2))
                                {

                                    using (Pen borderPen = new Pen(this.GridColor))
                                    {

                                        Rectangle borderDimensions = e.CellBounds;
                                        borderDimensions.Width -= 1;
                                        borderDimensions.Height -= 1;
                                        borderDimensions.X = borderDimensions.X + 1;
                                        borderDimensions.Y = borderDimensions.Y + 1;


                                        e.Graphics.DrawRectangle(eraser, borderDimensions);

                                        e.Graphics.FillRectangle(backBrush, borderDimensions);

                                        e.Graphics.DrawString(
                                            e.FormattedValue as string,
                                            this.Font,
                                            stringBrush,
                                            borderDimensions,
                                            new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far }
                                        );
                                        borderDimensions.Width += 1;
                                        borderDimensions.Height += 1;
                                        borderDimensions.X = borderDimensions.Left - 1;
                                        borderDimensions.Y = borderDimensions.Top - 1;

                                        e.Graphics.DrawRectangle(borderPen, borderDimensions);

                                        e.Handled = true;



                                    }
                                }
                            }
                            

                        }

                    }

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



            ((System.ComponentModel.ISupportInitialize)this).EndInit();

        }
    }
}
