using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spreadsheetApp
{
    internal class DataSource : DataTable
    {

        public DataSource(int rows, int cols) 
        {

            DataColumn Column;
            string columnName = "";
            for (int i = 0; i < cols; i++)
            {
                if (i <= 26) {
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(i + 65)}"; 
                } // 'A' is 65 in ASCII, so adding 64 to get A-Z.
                else
                {
                    // For columns beyond Z (i.e., AA, AB, etc.)
                    int quotient = (i - 1) / 26; // Calculate the "prefix" for double letters (A, B, etc.)
                    int remainder = (i - 1) % 26 + 1; // Calculate the "suffix" for double letters (A-Z)
                    // Combine the prefix and suffix to get AA, AB, etc.
                    Column = new DataColumn();
                    Column.ColumnName = $"{(char)(quotient + 65)}{(char)(remainder + 65)}";

                }
                Column.DataType = typeof(string);
                Column.AllowDBNull = true;
                Column.DefaultValue = "";
                Column.MaxLength = 255;
                this.Columns.Add(Column);
            }
            for (int j = 0; j < rows; j++)
            {
                this.Rows.Add(this.NewRow());
            }
        }
    }
}
