using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace spreadsheetApp
{
    public partial class PopUpForm : Form
    {
        public Document.DocParams Params {  get; set; }
   
        public PopUpForm(Document.DocParams docParams)
        {
            Params = docParams;
            InitializeComponent();
        }
        private void BtnCreateDoc_Click(object sender, EventArgs e)
        {

            try
            {
                Params.Title = docNameInput.Text;
                if (int.TryParse(numOfRowsInput.Text, out int rows)) Params.Rows = rows; else throw new Exception("Invalid number of rows.");
                if (int.TryParse(numOfColsInput.Text, out int cols)) Params.Columns = cols; else throw new Exception("Invalid number of columns.");
                DialogResult = DialogResult.OK;
                this.Close(); // popup is closed when document is created.
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message, "Invalid Parameters", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }
    }
}
