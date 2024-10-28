using System.Xml.Linq;

namespace spreadsheetApp
{

    public partial class SpreadsheetApp : Form
    {
        
        static int documentsCount;

        public string filePath;
        public Document CurrentFile { get; set; }
        public List<Document> Files { get; set; }
        private Document.DocParams Params { get; set; } 

        public SpreadsheetApp()
        {
            filePath = "C://";
            documentsCount = 0;
            InitializeComponent();
            Files = new List<Document>();   
            Params = new Document.DocParams();
        }
        private void _btnNew_Click(object sender, EventArgs e)
        {

            PopUpForm popup = new PopUpForm(Params); // to ask the user how many rows and columns and if he wants to name the sheet.
            if(popup.ShowDialog() == DialogResult.OK)
            {
                Document newDocument = new Document(Params.Title, Params.Rows, Params.Columns, filePath);
                Files.Add(newDocument);
                newDocument.Display();
                documentsCount++;
                this.Hide();// we cannot close this form because the app will be close, so we're hiding it.
            }
        }
        
        private void _btnOpen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.FileName = filePath;
                    ofd.Filter = "Excel | (*.xls)";
                    if(ofd.ShowDialog() == DialogResult.OK)
                    {
                        using (StreamReader sr = new StreamReader(ofd.FileName))
                        {
                            //Excel Libary goes Here
                            //File Manager Class Used
                            //FileManager Class could implement ISave and IOpen interface
                            //ISave saveFile() will be implemented by child classes
                            //JsonFileManager or XlsxFileManager implement ISave/IOpen
                            //Xlsx uses OpenXML or NPOI Library.
                        }
                    }
                }
            }
        }
    }
}
