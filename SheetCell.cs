﻿using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;

namespace spreadsheetApp
{
    [Serializable]
    public class SheetCell : DataGridViewTextBoxCell
    {
        private Rectangle cellArea;
        public Rectangle CellArea
        {
            get { return GetCellArea(); }
            set { cellArea = value; }
        }
        public System.Drawing.Color BackGroundColor { get; set; } = System.Drawing.Color.White;
        public System.Drawing.Color ForeColor { get; set; } = System.Drawing.SystemColors.InfoText;
        public System.Drawing.StringAlignment Alignment { get; set; } = System.Drawing.StringAlignment.Far;
        protected Rectangle GetCellArea() 
        {
            Rectangle r = new Rectangle();
            r.Width = cellArea.Width - 3;
            r.Height = cellArea.Height - 3;
            r.X = cellArea.X + 1;
            r.Y = cellArea.Y + 1;

            if (!this.Selected)
            {
                r.X = r.X - 1;
                r.Y = r.Y - 1;
                r.Width = r.Width + 2;
                r.Height = r.Height + 2; 
            }

            return r;
        }
        protected void DrawCellString(Graphics g, Rectangle cellBounds, string errorText)
        {
            if (string.IsNullOrEmpty(errorText)) 
            {
                CellArea = cellBounds;
                using (SolidBrush stringBrush = new SolidBrush(ForeColor))
                {
                    g.DrawString(
                        Value as string,
                        DataGridView.Font,
                        stringBrush,
                        CellArea,
                        new StringFormat { Alignment = Alignment, LineAlignment =Alignment }
                    );
                }
            }
            else
            {
                TextBox txtBox = this.DataGridView.EditingControl as TextBox;
                this.DataGridView.BeginEdit(true);
                txtBox.SelectionStart = int.Parse(errorText);
                txtBox.SelectionLength = 1;
                txtBox.Focus();
            }   
        }
        protected void FillCellBackGround(Graphics g, Rectangle cellBounds, System.Drawing.Color backGround)
        {
            CellArea = cellBounds;
            using (SolidBrush backBrush = new SolidBrush(backGround))
            {
                g.FillRectangle(backBrush, CellArea);
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

            if (this.RowIndex != -1 && this.ColumnIndex != -1) 
            {
                FillCellBackGround(graphics, cellBounds, BackGroundColor);
                DrawCellString(graphics, cellBounds, errorText);
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
                    int indexOfError = ife.InvalidChar;//Array.IndexOf(exp.ToCharArray(), ife.InvalidChar);
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
            string result;

            if (exp.StartsWith("=") || exp.StartsWith("+"))
            {
                char[] expArr = exp.ToCharArray();
                if (expArr[1] == 'A' && expArr[2]=='V' && expArr[3] == 'G')
                {
                    //DataGridViewSelectedCellCollection selectedCells = CurrentLa
                    result = "Sheet.AVG()";

                    return result;
                }
                else if(exp.Length > 2) 
                { 
                    exp = exp.Substring(1);
                    result = ExpressionProcessor.ProcessCommand(exp);
                    return result;
                }
                else return exp.Substring(1);
               
            }
            else return exp;
        }
        public override object Clone()
        {
            var clone = (SheetCell)base.Clone(); 
            clone.CellArea = CellArea;
            clone.BackGroundColor = BackGroundColor;
            clone.ForeColor = ForeColor;
            clone.Alignment = Alignment;
            return clone;
        }
        [Serializable]
        public class DataCell
        {
            public int CellRow {  get; set; }
            public int CellColumn { get; set; }
            public string CellValue { get; set; }
            public System.Drawing.Color BackGroundColor { get; set; }
            public System.Drawing.Color ForeColor { get; set; }
            public System.Drawing.StringAlignment Alignment { get; set; }
            public DataCell(int cellRow, int cellCol,string value, System.Drawing.Color backgroundColor, System.Drawing.Color foreColor, System.Drawing.StringAlignment alignment)
            {
                CellRow = cellRow;
                CellColumn = cellCol;
                CellValue = value;
                BackGroundColor = backgroundColor;
                ForeColor = foreColor;
                Alignment = alignment;
            }
        }
        [Serializable]
        public class DataCellSet
        {
            public List<DataCell> SelectedCells { get; set; } = new List<DataCell> ();

        }

    }
}
