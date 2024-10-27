namespace spreadsheetApp
{
    class SheetCell : DataGridViewTextBoxCell
    {
        public SheetCell()
        {
            Style.SelectionBackColor = Color.Wheat;
        }
        public void SetValue(object sender, DataGridViewCellEventArgs cell)
        {
            SheetCell editedCell = (SheetCell)DataGridView.Rows[cell.RowIndex].Cells[cell.ColumnIndex];
            string exp = (string)editedCell.Value;
            
            try
            {
                Value = ProccessExpression(exp);
            }
            catch (InvalidFormulaException ife)
            {
                if (MessageBox.Show(ife.Message, "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    
                    int indexOfError = Array.IndexOf(exp.ToCharArray(), ife.InvalidChar);
                    if (indexOfError != -1)
                    {
                        var cellTxt = this.AccessibilityObject;
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

            if (exp.StartsWith("="))
            {
                if (exp.Length > 1) exp = exp.Substring(1);
                return ExpressionProcessor.ProcessCommand(exp);
            }
            else return exp;
        }

    }

}
