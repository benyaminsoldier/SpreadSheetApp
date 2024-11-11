using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.CompilerServices;

namespace spreadsheetApp
{
     class SheetCell : DataGridViewTextBoxCell
    {
        private Rectangle cellArea;
        public bool IsFormatted {  get;  set; }  = false;
        public System.Drawing.Color BackGroundColor { get; set; } = System.Drawing.Color.White; 
        public Rectangle CellArea
        {
            get { return GetCellArea(); }
            set { cellArea = value; }
        }
        private Rectangle GetCellArea() 
        {
            Rectangle r = new Rectangle();
            r.Width = cellArea.Width - 1;
            r.Height = cellArea.Height - 1;
            r.X = cellArea.X + 1;
            r.Y = cellArea.Y + 1;
            return r;
        }
 
        public void DrawCellString(DataGridViewCellPaintingEventArgs e)
        {
            CellArea = e.CellBounds;
            using (SolidBrush stringBrush = new SolidBrush(this.DataGridView.ForeColor))
            {
                e.Graphics.DrawString(
                            e.FormattedValue as string,
                            this.DataGridView.Font,
                            stringBrush,
                            this.CellArea,
                            new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far }
                        );
            }
        }
        public void FillCellBackGround(DataGridViewCellPaintingEventArgs e)
        {
            CellArea = e.CellBounds;
            using (SolidBrush backBrush = new SolidBrush(System.Drawing.Color.White))
            {
                e.Graphics.FillRectangle(backBrush, CellArea);
            }

        }
        public void FillCellBackGround(System.Drawing.Color backgroundColor)
        {
                     
            Graphics g = this.DataGridView.CreateGraphics();
           // Rectangle cellBounds = this.DataGridView.CurrentCell.ContentBounds;
            //CellArea = cellBounds; 
            g.Clip = new Region(CellArea);
            using (SolidBrush backBrush = new SolidBrush(backgroundColor))
            {
                g.FillRectangle(backBrush, this.CellArea);
            }
            //this.DataGridView.Refresh();
        }
        public  void DrawNonSelectedBorders(DataGridViewCellPaintingEventArgs e)
        {           
            using (Pen eraser = new Pen(this.DataGridView.BackgroundColor, 2))
            {
                using (Pen borderPen = new Pen(this.DataGridView.GridColor))
                {
                    CellArea = e.CellBounds;
                    e.Graphics.DrawRectangle(eraser, CellArea);
                    FillCellBackGround(e);
                    DrawCellString(e);
                    e.Graphics.DrawRectangle(borderPen, e.CellBounds);
                    e.Handled = true;
                }
            }                       
        }
        public void DrawSelectedBorders(DataGridViewCellPaintingEventArgs e)
        {
            using (Pen borderpen = new Pen(System.Drawing.Color.DarkOliveGreen, 2))
            {
                
                CellArea = e.CellBounds;
                FillCellBackGround(e);
                DrawCellString(e);

                if (!String.IsNullOrEmpty(this.DataGridView.CurrentCell.ErrorText))
                {
                    TextBox txtBox = this.DataGridView.EditingControl as TextBox;
                    this.DataGridView.BeginEdit(true);
                    txtBox.SelectionStart = int.Parse(this.DataGridView.CurrentCell.ErrorText);
                    txtBox.SelectionLength = 1;
                    txtBox.Focus();
                }
                e.Graphics.DrawRectangle(borderpen, CellArea);
                e.Handled = true;

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
            clone.IsFormatted = IsFormatted;
            clone.CellArea = CellArea;
            clone.BackGroundColor = BackGroundColor;
            return clone;
        }

    }

}
