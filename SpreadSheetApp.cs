using System.Xml.Linq;

namespace spreadsheetApp
{
    public partial class SpreadsheetApp : Form
    {
        public string filePath;
        static int _documentsCount;

        public Document CurrentFile { get; set; }
        public List<Document> Files { get; set; }
        
        public SpreadsheetApp()
        {
            filePath = "C://";
            _documentsCount = 0;
            InitializeComponent();
            Files = new List<Document>();
        }
        private void _btnNew_Click(object sender, EventArgs e)
        {

            //PopupForm popup = new PopupForm(); // to ask the user how many rows and columns and if he wants to name the sheet.
            //popup.ShowDialog(); // This will block input to the main form until the popup is closed
            PopUpForm popup = new PopUpForm(); // to ask the user how many rows and columns and if he wants to name the sheet.
            popup.ShowDialog(); // this will block input to the main form until the popup is closed
            this.Hide(); // we cannot close this form because the app will be close, so we're hiding it.
            
            //Form popUpWindow = new Form(); 
            //popUpWindow.ShowDialog();
            //if(popOutWind.ShowDialog() == DialogResult.OK)
            //string name = "Doc1";
            //int numOfRows = 5;
            //int numOfColumns = 5;
            //Document newDocument = new Document(name, numOfRows, numOfColumns);

            //from the popout iwndow we take the name num of rows and columns and pass em to the Document constructo
            //So this way the popout window would just be collection the user infor and passin it again to the app.
            /*Document newDocument = new Document() {
                Name = "calculation_sheet" ,
                FilePath = filePath
            };
            Files.Add(newDocument);
            newDocument.Display();
            _documentsCount++;*/
            // filePath = ""; How to pass filePath to Document so when it saves the path gets updated.
            //Document newDocument = new Document() {
            //    Name = "calculation_sheet" ,
            //    FilePath = filePath
            //};
            //Files.Add(newDocument);
            //newDocument.Display();
            //// filePath = ""; How to pass filePath to Document so when it saves the path gets updated.
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
