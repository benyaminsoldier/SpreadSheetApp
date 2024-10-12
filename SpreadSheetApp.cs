namespace spreadsheetApp
{
    public partial class SpreadsheetApp : Form
    {
        string filePath;
        static int _documentsCount;
        public Document CurrentFile { get; set; }
        public List<Document> Files { get; set; }

        public SpreadsheetApp()
        {
            filePath = "C://";
            _documentsCount = 0;
            InitializeComponent();
            _documentsCount++;
            Files = new List<Document>();
        }



        private void _btnNew_Click(object sender, EventArgs e)
        {
            
            Document newDocument = new Document() {
                Name = "calculation_sheet" ,
                FilePath = filePath
            };
            Files.Add(newDocument);
            newDocument.Display();
            // filePath = ""; How to pass filePath to Document so when it saves the path gets updated.

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
