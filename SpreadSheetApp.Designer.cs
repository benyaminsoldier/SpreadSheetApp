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
            _menuBox.Location = new Point(8, 7);
            _menuBox.Margin = new Padding(2, 2, 2, 2);
            _menuBox.Name = "_menuBox";
            _menuBox.Padding = new Padding(2, 2, 2, 2);
            _menuBox.Size = new Size(210, 325);
            _menuBox.TabIndex = 0;
            _menuBox.TabStop = false;
            _menuBox.Text = "Menu";
            // 
            // _btnOpen
            // 
            _btnOpen.Location = new Point(42, 91);
            _btnOpen.Margin = new Padding(2, 2, 2, 2);
            _btnOpen.Name = "_btnOpen";
            _btnOpen.Size = new Size(78, 20);
            _btnOpen.TabIndex = 1;
            _btnOpen.Text = "Open File";
            _btnOpen.UseVisualStyleBackColor = true;
            _btnOpen.Click += _btnOpen_Click;
            // 
            // _btnNew
            // 
            _btnNew.Location = new Point(42, 59);
            _btnNew.Margin = new Padding(2, 2, 2, 2);
            _btnNew.Name = "_btnNew";
            _btnNew.Size = new Size(78, 20);
            _btnNew.TabIndex = 0;
            _btnNew.Text = "New File";
            _btnNew.UseVisualStyleBackColor = true;
            _btnNew.Click += _btnNew_Click;
            // 
            // SpreadsheetApp
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(899, 340);
            Controls.Add(_menuBox);
            Margin = new Padding(2, 2, 2, 2);
            Name = "SpreadsheetApp";
            Text = "Spreadsheet App";
            _menuBox.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private GroupBox _menuBox;
        private Button _btnOpen;
        private Button _btnNew;
    }
}
