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
            toolStripSplitButton1 = new ToolStripButton();
            toolBar.SuspendLayout();
            menuBar.SuspendLayout();
            flowLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // toolBar
            // 
            toolBar.ImageScalingSize = new Size(24, 24);
            toolBar.Items.AddRange(new ToolStripItem[] { pasteBtn, toolStripSeparator1, copyBtn, toolStripSeparator2, cutBtn, toolStripSeparator3, toolStripSplitButton1 });
            toolBar.Location = new Point(0, 33);
            toolBar.Name = "toolBar";
            toolBar.Size = new Size(1174, 33);
            toolBar.TabIndex = 9;
            toolBar.Text = "toolBar";
            // 
            // pasteBtn
            // 
            pasteBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            pasteBtn.Image = (Image)resources.GetObject("pasteBtn.Image");
            pasteBtn.ImageTransparentColor = Color.Magenta;
            pasteBtn.Name = "pasteBtn";
            pasteBtn.Size = new Size(34, 28);
            pasteBtn.Text = "pasteBtn";
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(6, 33);
            // 
            // copyBtn
            // 
            copyBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            copyBtn.Image = (Image)resources.GetObject("copyBtn.Image");
            copyBtn.ImageTransparentColor = Color.Magenta;
            copyBtn.Name = "copyBtn";
            copyBtn.Size = new Size(34, 28);
            copyBtn.Text = "copyBtn";
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new Size(6, 33);
            // 
            // cutBtn
            // 
            cutBtn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            cutBtn.Image = (Image)resources.GetObject("cutBtn.Image");
            cutBtn.ImageTransparentColor = Color.Magenta;
            cutBtn.Name = "cutBtn";
            cutBtn.Size = new Size(34, 28);
            cutBtn.Text = "cutBtn";
            // 
            // toolStripSeparator3
            // 
            toolStripSeparator3.Name = "toolStripSeparator3";
            toolStripSeparator3.Size = new Size(6, 33);
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
            // 
            // toolStripSplitButton1
            // 
            toolStripSplitButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripSplitButton1.Image = (Image)resources.GetObject("toolStripSplitButton1.Image");
            toolStripSplitButton1.ImageTransparentColor = Color.Magenta;
            toolStripSplitButton1.Name = "toolStripSplitButton1";
            toolStripSplitButton1.Size = new Size(34, 28);
            toolStripSplitButton1.Text = "toolStripSplitButton1";
            toolStripSplitButton1.Click += toolStripSplitButton1_Click;
            // 
            // Document
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1174, 562);
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
        private ColorsPallette colorsPallette1;
        private ColorsPallette colorsPallette2;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripButton toolStripSplitButton1;
    }
}