namespace spreadsheetApp
{
    class SheetCell : DataGridViewTextBoxCell
    {
        public SheetCell()
        {
            Style.SelectionBackColor = Color.Wheat;
        }
        public void SetValue(object sender, DataGridViewCellValidatingEventArgs cell)
        {
            
            string exp = cell.FormattedValue as string;
            
            try
            {
                Value = ProccessExpression(exp);
                this.DataGridView.InvalidateCell(this);
            }
            catch (InvalidFormulaException ife)
            {
                if (MessageBox.Show(ife.Message, "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    
                    int indexOfError = Array.IndexOf(exp.ToCharArray(), ife.InvalidChar);

                    if (indexOfError != -1)
                    {
                        var editedCell = this.DataGridView.CurrentCell;
                        this.DataGridView.BeginEdit(true); //Allows us to programmatically edit the cell textbox.
                        TextBox txtBox = this.DataGridView.EditingControl as TextBox;
                        txtBox.SelectionStart = indexOfError;
                        txtBox.SelectionLength = 1;
                        txtBox.Focus();
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

    }

}
