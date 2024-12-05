using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace spreadsheetApp
{
    public partial class ColorsPallette : UserControl
    {       
        public System.Drawing.Color CurrentColor {  get; set; }
        public List<System.Drawing.Color> Colors { get; set; }

        public delegate void ColorChosenEventHandler(object sender, ColorChosenEventArgs e);

        public event ColorChosenEventHandler ChosenColor;
        protected virtual void OnColorChosen(ColorChosenEventArgs e)
        {
            ChosenColor?.Invoke(this, e);
        }
        public ColorsPallette()
        {
            InitializeComponent();
            Colors = GetDefaultColorsList();
            PopulatePallette();
            Visible = false;
        }

        public ColorsPallette(List<System.Drawing.Color> colors)
        {
            InitializeComponent();
            Colors = colors.Take(10).ToList();
            PopulatePallette();
            Visible = false;
        }
        private List<System.Drawing.Color> GetDefaultColorsList()
        {

            KnownColor[] knowncolors = Enum.GetValues(typeof(KnownColor)) as KnownColor[];

            List<System.Drawing.Color> colors = new List<System.Drawing.Color>();
            List<System.Drawing.Color> found = new List<System.Drawing.Color>();
            List<System.Drawing.Color> temp = new List<System.Drawing.Color>();

            foreach (KnownColor knowncolor in knowncolors)
            {
                if (System.Drawing.Color.FromKnownColor(knowncolor) is System.Drawing.Color color && !color.IsSystemColor) temp.Add(color);
            }

            found = temp.Where(color => color.Name == "Black").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "White").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Red").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "DarkBlue").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Blue").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Orange").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Gray").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Yellow").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "LightBlue").ToList(); if (found.Count > 0) colors.Add(found[0]);
            found = temp.Where(color => color.Name == "Green").ToList(); if (found.Count > 0) colors.Add(found[0]);

            return colors;
        }
        private void PopulatePallette()
        {

            for (int col = 0; col < this.BackGroundScalesBase.ColumnCount; col++)
            {
                Panel nbp = new Panel()
                {
                    Name = $"{this.Colors[col].Name} Base",
                    BackColor = this.Colors[col],
                    Dock = DockStyle.Fill,
                    Margin = new Padding(2, 0, 2, 0),
                };
                nbp.Click += (sender, e) => 
                {
                    Document doc = this.Parent as Document;
                    Sheet grid = doc.CurrentLayout as Sheet;
                    SheetCell cell = grid.CurrentCell as SheetCell;

                    CurrentColor = nbp.BackColor;
                    OnColorChosen(new ColorChosenEventArgs() { ChosenColor = nbp.BackColor });
                    this.SendToBack();
                    this.Visible = false;
                };
                this.BackGroundScalesBase.Controls.Add(nbp, col, 0);
                for (int row = 0; row < this.BackGroundScales.RowCount; row++)
                {

                    Panel np = new Panel()
                    {
                        Name = $"{this.Colors[col].Name} Intensity({row * 10}%)",
                        BackColor = ControlPaint.Dark(this.Colors[col], ((row + 1) / 7.5f)),
                        Dock = DockStyle.Fill,
                        Margin = new Padding(2, 0, 2, 0),

                    };
                    if (nbp.BackColor == System.Drawing.Color.Black) np.BackColor = ControlPaint.Light(this.Colors[col], ((row + 1) / 7.5f));
                    np.Click += (sender, e) =>
                    {
                        
                        Document doc = this.Parent as Document;
                        Sheet grid = doc.CurrentLayout as Sheet;
                        SheetCell cell = grid.CurrentCell as SheetCell;
                        
                        CurrentColor = np.BackColor;
                        OnColorChosen(new ColorChosenEventArgs() { ChosenColor = np.BackColor});
                        this.SendToBack();
                        this.Visible = false;

                    };
                    np.Leave += (sender, e) => { };
                    np.MouseHover += (sender, e) => { };

                    this.BackGroundScales.Controls.Add(np, col, row);
                }
            }
        }
        private void btn_moreColors_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if(cd.ShowDialog() == DialogResult.OK)
                {
                    CurrentColor = cd.Color;
                    OnColorChosen(new ColorChosenEventArgs() { ChosenColor = cd.Color });
                    this.SendToBack();
                    this.Visible = false;
                }                
            }
        }
    }
}
