namespace spreadsheetApp
{
    class SheetCell : DataGridViewTextBoxCell
    {

        public void SetValue(object sender, DataGridViewCellEventArgs e) 
        {
            SheetCell editedCell = (SheetCell)DataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string exp = (string)editedCell.Value;
            Console.WriteLine(exp);
            Value = ProccessExpression(exp);
        }

        public string ProccessExpression(string exp) 
        {            
            return "default"; 
        }

    }

}
