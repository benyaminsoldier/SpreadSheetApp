using spreadsheetApp;

class SheetColumn : DataGridViewTextBoxColumn
{
    public SheetColumn() 
    { 
        this.SortMode = DataGridViewColumnSortMode.NotSortable;
        this.Width = 150;
        this.CellTemplate = new SheetCell();
    }

}
