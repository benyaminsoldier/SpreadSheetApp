using System.Xml.Linq;

namespace spreadsheetApp
{
    public partial class SpreadsheetApp : Form
    {
        string filePath;
        static int _documentsCount;
        public Document Document { get; set; }
        public List<Document> Documents { get; set; }
        public SpreadsheetApp()
        {
            InitializeComponent();
            _documentsCount++;
            Documents = new List<Document>();
        }
        private void _btnNew_Click(object sender, EventArgs e)
        {
            //PopupForm popup = new PopupForm(); // to ask the user how many rows and columns and if he wants to name the sheet.
            //popup.ShowDialog(); // This will block input to the main form until the popup is closed
            
            //Form popUpWindow = new Form(); 
            //popUpWindow.ShowDialog();
            //string name = "Doc1";
            //int numOfRows = 5;
            //int numOfColumns = 5;
            //Document newDocument = new Document(name, numOfRows, numOfColumns);
            Document newDocument = new Document();
            Documents.Add(newDocument);
            newDocument.Display();
            filePath = "";
        }
        private void _btnOpen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.FileName = "C://";
                    ofd.Filter = "Excel | (*.xls)";
                    if(ofd.ShowDialog() == DialogResult.OK)
                    {
                        using (StreamReader sr = new StreamReader(ofd.FileName))
                        {
                            //sr.ReadToEnd(); Pending for research 
                        }
                    }
                }
            }
        }
    }
}
