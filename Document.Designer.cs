namespace spreadsheetApp
{
    partial class Document
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Document));
            toolBar = new ToolStrip();
            pasteBtn = new ToolStripButton();
            toolStripSeparator1 = new ToolStripSeparator();
            copyBtn = new ToolStripButton();
            toolStripSeparator2 = new ToolStripSeparator();
            cutBtn = new ToolStripButton();
            toolStripSeparator3 = new ToolStripSeparator();
            btn_backGroundCellFormat = new ToolStripButton();
            toolStripSeparator4 = new ToolStripSeparator();
            btn_fontColor = new ToolStripButton();
            toolStripSeparator5 = new ToolStripSeparator();
            btn_leftAlign = new ToolStripButton();
            btn_centerAlign = new ToolStripButton();
            btn_rightAlign = new ToolStripButton();
            toolStripSeparator6 = new ToolStripSeparator();
            menuBar = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            newToolStripMenuItem = new ToolStripMenuItem();
            openToolStripMenuItem = new ToolStripMenuItem();
            saveToolStripMenuItem = new ToolStripMenuItem();
            saveAsToolStripMenuItem = new ToolStripMenuItem();
            printToolStripMenuItem = new ToolStripMenuItem();
            exitToolStripMenuItem = new ToolStripMenuItem();
            flowLayoutPanel1 = new FlowLayoutPanel();
            btnSheet = new Button();
            colorsPallette2 = new ColorsPallette();
            splitContainer1 = new SplitContainer();
            colorsPallette1 = new ColorsPallette();
            toolStripButton1 = new ToolStripDropDownButton();
            avgMenuItem = new ToolStripMenuItem();
            sumMenuItem = new ToolStripMenuItem();
            toolBar.SuspendLayout();
            menuBar.SuspendLayout();
            flowLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // toolBar
            // 
            toolBar.ImageScalingSize = new Size(24, 24);
            toolBar.Items.AddRange(new ToolStripItem[] { pasteBtn, toolStripSeparator1, copyBtn, toolStripSeparator2, cutBtn, toolStripSeparator3, btn_backGroundCellFormat, toolStripSeparator4, btn_fontColor, toolStripSeparator5, btn_leftAlign, btn_centerAlign, btn_rightAlign, toolStripSeparator6, toolStripButton1 });
            toolBar.Location = new Point(0, 33);
            toolBar.Name = "toolBar";
            toolBar.Size = new Size(1174, 35);
            toolBar.TabIndex = 9;
            toolBar.Text = "toolBar";
            // 
            // pasteBtn
            // 
            pasteBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            pasteBtn.Image = Properties.Resources.pasteIcon;
            pasteBtn.ImageTransparentColor = Color.Magenta;
            pasteBtn.Name = "pasteBtn";
            pasteBtn.Size = new Size(34, 30);
            pasteBtn.Text = "pasteBtn";
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(6, 35);
            // 
            // copyBtn
            // 
            copyBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            copyBtn.Image = Properties.Resources.copyIcon;
            copyBtn.ImageTransparentColor = Color.Magenta;
            copyBtn.Name = "copyBtn";
            copyBtn.Size = new Size(34, 30);
            copyBtn.Text = "copyBtn";
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new Size(6, 35);
            // 
            // cutBtn
            // 
            cutBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            cutBtn.Image = Properties.Resources.cutIcon;
            cutBtn.ImageTransparentColor = Color.Magenta;
            cutBtn.Name = "cutBtn";
            cutBtn.Size = new Size(34, 30);
            cutBtn.Text = "cutBtn";
            // 
            // toolStripSeparator3
            // 
            toolStripSeparator3.Name = "toolStripSeparator3";
            toolStripSeparator3.Size = new Size(6, 35);
            // 
            // btn_backGroundCellFormat
            // 
            btn_backGroundCellFormat.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btn_backGroundCellFormat.Image = Properties.Resources.backgroundColorIcon;
            btn_backGroundCellFormat.ImageTransparentColor = Color.Magenta;
            btn_backGroundCellFormat.Name = "btn_backGroundCellFormat";
            btn_backGroundCellFormat.Size = new Size(34, 30);
            btn_backGroundCellFormat.Text = "toolStripSplitButton1";
            btn_backGroundCellFormat.Click += BackGroundCellFormatBtn_Click;
            // 
            // toolStripSeparator4
            // 
            toolStripSeparator4.Name = "toolStripSeparator4";
            toolStripSeparator4.Size = new Size(6, 35);
            // 
            // btn_fontColor
            // 
            btn_fontColor.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btn_fontColor.Image = Properties.Resources.fontColor3;
            btn_fontColor.ImageTransparentColor = Color.Magenta;
            btn_fontColor.Name = "btn_fontColor";
            btn_fontColor.Size = new Size(34, 30);
            btn_fontColor.Click += btn_fontColor_Click;
            // 
            // toolStripSeparator5
            // 
            toolStripSeparator5.Name = "toolStripSeparator5";
            toolStripSeparator5.Size = new Size(6, 35);
            // 
            // btn_leftAlign
            // 
            btn_leftAlign.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btn_leftAlign.Image = Properties.Resources.leftAlignIcon4;
            btn_leftAlign.ImageTransparentColor = Color.Magenta;
            btn_leftAlign.Name = "btn_leftAlign";
            btn_leftAlign.Size = new Size(34, 30);
            btn_leftAlign.Click += btn_leftAlign_Click;
            // 
            // btn_centerAlign
            // 
            btn_centerAlign.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btn_centerAlign.Image = Properties.Resources.centerAlignIcon1;
            btn_centerAlign.ImageTransparentColor = Color.Magenta;
            btn_centerAlign.Name = "btn_centerAlign";
            btn_centerAlign.Size = new Size(34, 30);
            btn_centerAlign.Text = "toolStripButton1";
            btn_centerAlign.Click += btn_centerAlign_Click;
            // 
            // btn_rightAlign
            // 
            btn_rightAlign.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btn_rightAlign.Image = Properties.Resources.rightAlignIcon1;
            btn_rightAlign.ImageTransparentColor = Color.Magenta;
            btn_rightAlign.Name = "btn_rightAlign";
            btn_rightAlign.Padding = new Padding(1);
            btn_rightAlign.Size = new Size(34, 30);
            btn_rightAlign.Text = "toolStripButton2";
            btn_rightAlign.Click += btn_rightAlign_Click;
            // 
            // toolStripSeparator6
            // 
            toolStripSeparator6.Name = "toolStripSeparator6";
            toolStripSeparator6.Size = new Size(6, 35);
            // 
            // menuBar
            // 
            menuBar.ImageScalingSize = new Size(24, 24);
            menuBar.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
            menuBar.Location = new Point(0, 0);
            menuBar.Name = "menuBar";
            menuBar.Padding = new Padding(8, 2, 0, 2);
            menuBar.Size = new Size(1174, 33);
            menuBar.TabIndex = 8;
            menuBar.Text = "menuBar";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { newToolStripMenuItem, openToolStripMenuItem, saveToolStripMenuItem, saveAsToolStripMenuItem, printToolStripMenuItem, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(54, 29);
            fileToolStripMenuItem.Text = "File";
            // 
            // newToolStripMenuItem
            // 
            newToolStripMenuItem.Name = "newToolStripMenuItem";
            newToolStripMenuItem.Size = new Size(188, 34);
            newToolStripMenuItem.Text = "New";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new Size(188, 34);
            openToolStripMenuItem.Text = "Open";
            // 
            // saveToolStripMenuItem
            // 
            saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            saveToolStripMenuItem.Size = new Size(188, 34);
            saveToolStripMenuItem.Text = "Save";
            // 
            // saveAsToolStripMenuItem
            // 
            saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            saveAsToolStripMenuItem.Size = new Size(188, 34);
            saveAsToolStripMenuItem.Text = "Save As...";
            saveAsToolStripMenuItem.Click += saveAsToolStripMenuItem_Click;
            // 
            // printToolStripMenuItem
            // 
            printToolStripMenuItem.Name = "printToolStripMenuItem";
            printToolStripMenuItem.Size = new Size(188, 34);
            printToolStripMenuItem.Text = "Print";
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new Size(188, 34);
            exitToolStripMenuItem.Text = "Exit";
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(btnSheet);
            flowLayoutPanel1.Dock = DockStyle.Bottom;
            flowLayoutPanel1.Location = new Point(0, 504);
            flowLayoutPanel1.Margin = new Padding(4);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(1174, 58);
            flowLayoutPanel1.TabIndex = 15;
            // 
            // btnSheet
            // 
            btnSheet.AutoSize = true;
            btnSheet.Location = new Point(4, 4);
            btnSheet.Margin = new Padding(4);
            btnSheet.Name = "btnSheet";
            btnSheet.Size = new Size(186, 44);
            btnSheet.TabIndex = 4;
            btnSheet.Text = "Sheet 1";
            btnSheet.UseVisualStyleBackColor = true;
            // 
            // colorsPallette2
            // 
            colorsPallette2.BackColor = SystemColors.ControlLightLight;
            colorsPallette2.Colors = (List<Color>)resources.GetObject("colorsPallette2.Colors");
            colorsPallette2.CurrentColor = Color.Empty;
            colorsPallette2.Location = new Point(133, 66);
            colorsPallette2.Margin = new Padding(0);
            colorsPallette2.Name = "colorsPallette2";
            colorsPallette2.Padding = new Padding(1);
            colorsPallette2.Size = new Size(273, 238);
            colorsPallette2.TabIndex = 16;
            colorsPallette2.Visible = false;
            // 
            // splitContainer1
            // 
            splitContainer1.Location = new Point(293, 418);
            splitContainer1.Name = "splitContainer1";
            splitContainer1.Size = new Size(225, 150);
            splitContainer1.SplitterDistance = 75;
            splitContainer1.TabIndex = 17;
            // 
            // colorsPallette1
            // 
            colorsPallette1.BackColor = SystemColors.ControlLightLight;
            colorsPallette1.Colors = (List<Color>)resources.GetObject("colorsPallette1.Colors");
            colorsPallette1.CurrentColor = Color.Empty;
            colorsPallette1.Location = new Point(172, 66);
            colorsPallette1.Margin = new Padding(0);
            colorsPallette1.Name = "colorsPallette1";
            colorsPallette1.Padding = new Padding(1);
            colorsPallette1.Size = new Size(273, 238);
            colorsPallette1.TabIndex = 18;
            colorsPallette1.Visible = false;
            // 
            // toolStripButton1
            // 
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButton1.DropDownItems.AddRange(new ToolStripItem[] { avgMenuItem, sumMenuItem });
            toolStripButton1.Image = Properties.Resources.opsIcons1;
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(42, 30);
            toolStripButton1.Text = "toolStripButton1";
            // 
            // avgMenuItem
            // 
            avgMenuItem.Name = "avgMenuItem";
            avgMenuItem.Size = new Size(270, 34);
            avgMenuItem.Text = "AVG";
            avgMenuItem.Click += avgMenuItem_Click;
            // 
            // sumMenuItem
            // 
            sumMenuItem.Name = "sumMenuItem";
            sumMenuItem.Size = new Size(270, 34);
            sumMenuItem.Text = "SUM";
            sumMenuItem.Click += sumMenuItem_Click;
            // 
            // Document
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1174, 562);
            Controls.Add(colorsPallette1);
            Controls.Add(splitContainer1);
            Controls.Add(colorsPallette2);
            Controls.Add(flowLayoutPanel1);
            Controls.Add(toolBar);
            Controls.Add(menuBar);
            Margin = new Padding(4);
            Name = "Document";
            Text = "Doc";
            Load += Document_Load;
            toolBar.ResumeLayout(false);
            toolBar.PerformLayout();
            menuBar.ResumeLayout(false);
            menuBar.PerformLayout();
            flowLayoutPanel1.ResumeLayout(false);
            flowLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private ToolStrip toolBar;
        private ToolStripButton pasteBtn;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton copyBtn;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton cutBtn;
        private MenuStrip menuBar;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem newToolStripMenuItem;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripMenuItem saveToolStripMenuItem;
        private ToolStripMenuItem saveAsToolStripMenuItem;
        private ToolStripMenuItem printToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private FlowLayoutPanel flowLayoutPanel1;
        private Button btnSheet;
        private ColorsPallette colorsPallette2;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripButton btn_backGroundCellFormat;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripButton btn_fontColor;
        private SplitContainer splitContainer1;
        private ColorsPallette colorsPallette1;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripButton btn_leftAlign;
        private ToolStripButton btn_centerAlign;
        private ToolStripButton btn_rightAlign;
        private ToolStripSeparator toolStripSeparator6;
        private ToolStripDropDownButton toolStripButton1;
        private ToolStripMenuItem avgMenuItem;
        private ToolStripMenuItem sumMenuItem;
    }
}