namespace spreadsheetApp
{
    partial class ColorsPallette
    {
        /// <summary> 
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary> 
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            BackGroundScales = new TableLayoutPanel();
            header_pallette = new GroupBox();
            label1 = new Label();
            BackGroundScalesBase = new TableLayoutPanel();
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            btn_moreColors = new Button();
            header_pallette.SuspendLayout();
            SuspendLayout();
            // 
            // BackGroundScales
            // 
            BackGroundScales.ColumnCount = 10;
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.11111F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 11.1111107F));
            BackGroundScales.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 25F));
            BackGroundScales.Location = new Point(4, 69);
            BackGroundScales.Margin = new Padding(0);
            BackGroundScales.Name = "BackGroundScales";
            BackGroundScales.RowCount = 5;
            BackGroundScales.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            BackGroundScales.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            BackGroundScales.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            BackGroundScales.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            BackGroundScales.RowStyles.Add(new RowStyle(SizeType.Percent, 20F));
            BackGroundScales.Size = new Size(264, 133);
            BackGroundScales.TabIndex = 0;
            // 
            // header_pallette
            // 
            header_pallette.BackColor = SystemColors.ControlLight;
            header_pallette.Controls.Add(label1);
            header_pallette.Dock = DockStyle.Top;
            header_pallette.Location = new Point(1, 1);
            header_pallette.Name = "header_pallette";
            header_pallette.Size = new Size(273, 32);
            header_pallette.TabIndex = 1;
            header_pallette.TabStop = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = SystemColors.ControlLight;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.ForeColor = SystemColors.ControlText;
            label1.Location = new Point(74, 0);
            label1.Name = "label1";
            label1.Size = new Size(120, 25);
            label1.TabIndex = 0;
            label1.Text = "Theme Color";
            // 
            // BackGroundScalesBase
            // 
            BackGroundScalesBase.ColumnCount = 10;
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            BackGroundScalesBase.Location = new Point(4, 37);
            BackGroundScalesBase.Name = "BackGroundScalesBase";
            BackGroundScalesBase.RowCount = 1;
            BackGroundScalesBase.RowStyles.Add(new RowStyle());
            BackGroundScalesBase.Size = new Size(263, 25);
            BackGroundScalesBase.TabIndex = 2;
            // 
            // btn_moreColors
            // 
            btn_moreColors.BackColor = SystemColors.ControlLight;
            btn_moreColors.Dock = DockStyle.Bottom;
            btn_moreColors.Location = new Point(1, 203);
            btn_moreColors.Name = "btn_moreColors";
            btn_moreColors.Size = new Size(273, 34);
            btn_moreColors.TabIndex = 3;
            btn_moreColors.Text = "More Colors...";
            btn_moreColors.UseVisualStyleBackColor = false;
            btn_moreColors.Click += btn_moreColors_Click;
            // 
            // ColorsPallette
            // 
            Controls.Add(btn_moreColors);
            Controls.Add(BackGroundScalesBase);
            Controls.Add(header_pallette);
            Controls.Add(BackGroundScales);
            Margin = new Padding(0);
            Name = "ColorsPallette";
            Padding = new Padding(1);
            Size = new Size(275, 238);
            header_pallette.ResumeLayout(false);
            header_pallette.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel BackGroundScales;
        private GroupBox header_pallette;
        private Label label1;
        private TableLayoutPanel BackGroundScalesBase;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Button btn_moreColors;
    }
}
