using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        public int NumOfRows { get; set; } = 0;
        public int NumOfCols { get; set; } = 0;
        
        public SpreadsheetApp()
        {
            filePath = "";
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
                //this.Hide();// we cannot close this form because the app will be close, so we're hiding it.
            }            
        }
        private void _btnOpen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.FileName = "";
                    //ofd.Filter = "Excel | (*.xlsx)";
                    if (ofd.ShowDialog() == DialogResult.OK)

                    {
                        //using (StreamReader sr = new StreamReader(ofd.FileName))
                        // I THINK WE DONOT NEED STREAMREADER, RIGHT? WE'RE NOT READING NOTEPAD FILES.

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
                            } else
                            {
                                NumOfCols = 0;
                            }
                            if (NumOfCols == 0 && NumOfRows == 0)
                            {
                                Document newOpenedDocument = new Document(ofd.FileName, 50, 50, filePath);
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

                        // USING NPOI PACKAGE
                        //File Manager Class Used
                        //FileManager Class could implement ISave and IOpen interface
                        //ISave saveFile() will be implemented by child classes
                        //JsonFileManager or XlsxFileManager implement ISave/IOpen
                        //Xlsx uses OpenXML or NPOI Library.
                        //}
                    }
                }
            }
        }
    }
}
