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
        // just copied this property to the PopUp file because we're creating the file from this form.
        // maybe there's a better way of doing?? but I just copied.
        public Document CurrentFile { get; set; }
        public List<Document> Files { get; set; } = new List<Document>();
        public PopUpForm()
        {
            InitializeComponent();
        }
        private void BtnCreateDoc_Click(object sender, EventArgs e)
        {
            string docName = docNameInput.Text;
            int rows = int.Parse(numOfRowsInput.Text);
            int cols = int.Parse(numOfColsInput.Text);
            string filePath = "C://";
            Document newDocument = new Document(docName, rows, cols, filePath);
            Files.Add(newDocument);
            newDocument.Display();
            this.Close(); // popup is closed when document is created.
            // filePath = ""; How to pass filePath to Document so when it saves the path gets updated.
        }
    }
}
