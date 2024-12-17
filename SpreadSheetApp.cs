using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Windows.Forms;
using System;
using System.Linq;


namespace spreadsheetApp
{ 
    public partial class SpreadsheetApp : Form
    { 
        static int documentsCount;
        public string filePath;
        public Document CurrentFile { get; set; }
        public List<Document> Files { get; set; }
        private Document.DocParams Params { get; set; } 
        public int NumOfRows { get; set; } = 0;
        public int NumOfCols { get; set; } = 0;

        public SpreadsheetApp()
        {
            filePath = "";
            documentsCount = 0;
            InitializeComponent();
            Files = new List<Document>();   
            Params = new Document.DocParams();
            this.FormClosing += MainForm_FormClosing; // handling first window being closed.
        }
        private void _btnNew_Click(object sender, EventArgs e)
        {
            PopUpForm popup = new PopUpForm(Params); // to ask the user how many rows and columns and if he wants to name the sheet.
            if(popup.ShowDialog() == DialogResult.OK)
            {
                Document newDocument = new Document(Params.Title, Params.Rows, Params.Columns, filePath);
                newDocument.Display();
                documentsCount++;
            }            
        }
        private void _btnOpen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Open File";
                    ofd.FileName = "";
                    ofd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                    filePath = ofd.FileName;
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        // USING OPENXML
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ofd.FileName, false))
                        {
                            WorkbookPart workbookPart = doc.WorkbookPart;
                            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet1 = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First();
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet1.Id);
                            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                            NumOfRows = worksheetPart.Worksheet.Descendants<Row>().Count();
                            Row FirstRow = worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault();
                            if (FirstRow != null)
                            {
                                NumOfCols = FirstRow.Descendants<Cell>().Count();
                            }
                            else
                            {
                                NumOfCols = 0;
                            }
                            if (NumOfCols == 0 && NumOfRows == 0)
                            {
                                Document newOpenedDocument = new Document(ofd.FileName, 30, 10, filePath);
                                Files.Add(newOpenedDocument);
                                newOpenedDocument.Display();
                            }
                            else
                            {
                                Document newOpenedDocument = new Document(ofd.FileName, NumOfRows, NumOfCols, filePath, doc); 
                                Files.Add(newOpenedDocument);
                                newOpenedDocument.Display();
                            }
                        }
                    }
                }
            }
        }
        
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Application.OpenForms.Count > 1) // if the other forms are open, we just hide the initial form.
            {
                this.Hide();
                e.Cancel = true;
            }
            else
            {
                Application.Exit(); // if it's the last one open, close the application.
            }
        }
    }
}
