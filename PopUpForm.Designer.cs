namespace spreadsheetApp
{
    partial class PopUpForm
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
            docNameInput = new TextBox();
            numOfRowsInput = new TextBox();
            numOfColsInput = new TextBox();
            BtnCreateDoc = new Button();
            docNameLabel = new Label();
            numOfRowsLabel = new Label();
            numOfColsLabel = new Label();
            SuspendLayout();
            // 
            // docNameInput
            // 
            docNameInput.Location = new Point(388, 28);
            docNameInput.Name = "docNameInput";
            docNameInput.Size = new Size(125, 27);
            docNameInput.TabIndex = 0;
            // 
            // numOfRowsInput
            // 
            numOfRowsInput.Location = new Point(169, 80);
            numOfRowsInput.Name = "numOfRowsInput";
            numOfRowsInput.Size = new Size(125, 27);
            numOfRowsInput.TabIndex = 1;
            // 
            // numOfColsInput
            // 
            numOfColsInput.Location = new Point(200, 131);
            numOfColsInput.Name = "numOfColsInput";
            numOfColsInput.Size = new Size(125, 27);
            numOfColsInput.TabIndex = 2;
            // 
            // BtnCreateDoc
            // 
            BtnCreateDoc.Location = new Point(237, 188);
            BtnCreateDoc.Name = "BtnCreateDoc";
            BtnCreateDoc.Size = new Size(108, 39);
            BtnCreateDoc.TabIndex = 3;
            BtnCreateDoc.Text = "Create";
            BtnCreateDoc.UseVisualStyleBackColor = true;
            BtnCreateDoc.Click += BtnCreateDoc_Click;
            // 
            // docNameLabel
            // 
            docNameLabel.AutoSize = true;
            docNameLabel.Font = new Font("Segoe UI", 10F);
            docNameLabel.Location = new Point(23, 31);
            docNameLabel.Name = "docNameLabel";
            docNameLabel.Size = new Size(357, 23);
            docNameLabel.TabIndex = 4;
            docNameLabel.Text = "How would you like to name your document?";
            // 
            // numOfRowsLabel
            // 
            numOfRowsLabel.AutoSize = true;
            numOfRowsLabel.Font = new Font("Segoe UI", 10F);
            numOfRowsLabel.Location = new Point(23, 83);
            numOfRowsLabel.Name = "numOfRowsLabel";
            numOfRowsLabel.Size = new Size(137, 23);
            numOfRowsLabel.TabIndex = 5;
            numOfRowsLabel.Text = "Number of rows:";
            // 
            // numOfColsLabel
            // 
            numOfColsLabel.AutoSize = true;
            numOfColsLabel.Font = new Font("Segoe UI", 10F);
            numOfColsLabel.Location = new Point(23, 134);
            numOfColsLabel.Name = "numOfColsLabel";
            numOfColsLabel.Size = new Size(169, 23);
            numOfColsLabel.TabIndex = 6;
            numOfColsLabel.Text = "Number of Columns:";
            // 
            // PopUpForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(582, 253);
            Controls.Add(numOfColsLabel);
            Controls.Add(numOfRowsLabel);
            Controls.Add(docNameLabel);
            Controls.Add(BtnCreateDoc);
            Controls.Add(numOfColsInput);
            Controls.Add(numOfRowsInput);
            Controls.Add(docNameInput);
            Name = "PopUpForm";
            Text = "Sheet Configuration";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox docNameInput;
        private TextBox numOfRowsInput;
        private TextBox numOfColsInput;
        private Button BtnCreateDoc;
        private Label docNameLabel;
        private Label numOfRowsLabel;
        private Label numOfColsLabel;
    }
}