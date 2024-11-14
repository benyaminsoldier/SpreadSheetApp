using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.CompilerServices;

namespace spreadsheetApp
{
     class SheetCell : DataGridViewTextBoxCell
    {
        private Rectangle cellArea;
        public Rectangle CellArea
        {
            get { return GetCellArea(); }
            set { cellArea = value; }
        }


        public System.Drawing.Color BackGroundColor { get; set; } = System.Drawing.Color.White; 

        protected Rectangle GetCellArea() 
        {
            Rectangle r = new Rectangle();
            r.Width = cellArea.Width - 3;
            r.Height = cellArea.Height - 3;
            r.X = cellArea.X + 1;
            r.Y = cellArea.Y + 1;
            return r;
        }
        protected void DrawCellString(Graphics g, Rectangle cellBounds)
        {
            CellArea = cellBounds;
            using (SolidBrush stringBrush = new SolidBrush(System.Drawing.Color.Black))
            {
                g.DrawString(
                            Value as string,
                            DataGridView.Font,
                            stringBrush,
                            CellArea,
                            new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far }
                        );
            }
        }
        protected void FillCellBackGround(Graphics g, Rectangle cellBounds, System.Drawing.Color backGround)
        {
            
            CellArea = cellBounds;
            if (!this.Selected)
            {
                Rectangle r = new Rectangle()
                {
                    X = CellArea.X-1,
                    Y = CellArea.Y-1,
                    Width = CellArea.Width +2,
                    Height = CellArea.Height + 2
                };
                using (SolidBrush backBrush = new SolidBrush(backGround))
                {
                    g.FillRectangle(backBrush, r);
                }
                ;
            }
            else
            {
                using (SolidBrush backBrush = new SolidBrush(backGround))
                {
                    g.FillRectangle(backBrush, CellArea);
                }
            }
            

        }
        protected void FillCellBackGround(System.Drawing.Color backgroundColor)
        {
                     
            Graphics g = this.DataGridView.CreateGraphics();
            g.Clip = new Region(CellArea);
            using (SolidBrush backBrush = new SolidBrush(backgroundColor))
            {
                g.FillRectangle(backBrush, CellArea);
            }
            
        }

        protected void DrawSelectedCellBorder(Graphics g, Rectangle cellBounds)
        {
            using (Pen borderpen = new Pen(System.Drawing.Color.DarkOliveGreen, 2))
            {
                CellArea = cellBounds;
                g.DrawRectangle(borderpen, CellArea);
            }
        }

        protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts)
        {
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);

            if (this.RowIndex >= 0 && this.ColumnIndex >= 0) 
            {
                FillCellBackGround(graphics, cellBounds, BackGroundColor);
                if (string.IsNullOrEmpty(errorText)) DrawCellString(graphics, cellBounds);
                else
                {
                    
                }

                
                if (this.Selected) DrawSelectedCellBorder(graphics, cellBounds);
            }
        }


        public void SetValue(object sender, DataGridViewCellValidatingEventArgs cell)
        {
            
            string exp = cell.FormattedValue as string;
            
            try
            {
                Value = ProccessExpression(exp);
                this.ErrorText = null;
            }
            catch (InvalidFormulaException ife)
            {
                if (MessageBox.Show(ife.Message, "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    
                    int indexOfError = exp.IndexOf(ife.InvalidChar);//Array.IndexOf(exp.ToCharArray(), ife.InvalidChar);
                    
                    if (indexOfError != -1) 
                    {
                        this.ErrorText = indexOfError.ToString();
                        this.DataGridView.BeginEdit(true); //Allows us to programmatically edit the cell textbox.
                        TextBox txtBox = this.DataGridView.EditingControl as TextBox;
                        txtBox.Text = exp;
                        this.Value = exp;
                        txtBox.SelectionStart = indexOfError;
                        txtBox.SelectionLength = 1;                        
                        cell.Cancel = true;
                    }
                }               
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        private string ProccessExpression(string exp)
        {
            if (string.IsNullOrEmpty(exp)) return exp;

            exp = exp.Trim();

            if (exp.StartsWith("=") || exp.StartsWith("+") || exp.StartsWith("-"))
            {
                if (exp.Length > 2) exp = exp.Substring(1);
                else return exp.Substring(1);
                return ExpressionProcessor.ProcessCommand(exp);               
            }
            else return exp;
        }
        public override object Clone()
        {
            var clone = (SheetCell)base.Clone(); 
            clone.CellArea = CellArea;
            clone.BackGroundColor = BackGroundColor;
            return clone;
        }

    }

}
