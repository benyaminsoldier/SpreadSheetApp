namespace spreadsheetApp
{
    public partial class SpreadsheetApp : Form
    {
        string filePath;
        static int _documentsCount;
        
        public List<Document> Documents { get; set; }

        public SpreadsheetApp()
        {
            InitializeComponent();
            _documentsCount++;
            Documents = new List<Document>();
        }



        private void _btnNew_Click(object sender, EventArgs e)
        {

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
