namespace spreadsheetApp
{
    partial class SpreadsheetApp
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            _menuBox = new GroupBox();
            _btnOpen = new Button();
            _btnNew = new Button();
            _menuBox.SuspendLayout();
            SuspendLayout();
            // 
            // _menuBox
            // 
            _menuBox.Controls.Add(_btnOpen);
            _menuBox.Controls.Add(_btnNew);
            _menuBox.Location = new Point(12, 12);
            _menuBox.Name = "_menuBox";
            _menuBox.Size = new Size(300, 542);
            _menuBox.TabIndex = 0;
            _menuBox.TabStop = false;
            _menuBox.Text = "Menu";
            // 
            // _btnOpen
            // 
            _btnOpen.Location = new Point(60, 152);
            _btnOpen.Name = "_btnOpen";
            _btnOpen.Size = new Size(112, 34);
            _btnOpen.TabIndex = 1;
            _btnOpen.Text = "Open File";
            _btnOpen.UseVisualStyleBackColor = true;
            _btnOpen.Click += _btnOpen_Click;
            // 
            // _btnNew
            // 
            _btnNew.Location = new Point(60, 98);
            _btnNew.Name = "_btnNew";
            _btnNew.Size = new Size(112, 34);
            _btnNew.TabIndex = 0;
            _btnNew.Text = "New File";
            _btnNew.UseVisualStyleBackColor = true;
            _btnNew.Click += _btnNew_Click;
            // 
            // SpreadsheetApp
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1284, 566);
            Controls.Add(_menuBox);
            Name = "SpreadsheetApp";
            Text = "Excel";
            _menuBox.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private GroupBox _menuBox;
        private Button _btnOpen;
        private Button _btnNew;
    }
}
