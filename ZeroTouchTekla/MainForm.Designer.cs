
namespace ZeroTouchTekla
{
    partial class MainForm
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
            this.SetExcelButton = new System.Windows.Forms.Button();
            this.FootingRebarButton = new System.Windows.Forms.Button();
            this.ExcelNameLabel = new System.Windows.Forms.Label();
            this.SheetComboBox = new System.Windows.Forms.ComboBox();
            this.LoadButton = new System.Windows.Forms.Button();
            this.CopySpacingButton = new System.Windows.Forms.Button();
            this.utilityPanel = new System.Windows.Forms.Panel();
            this.recreateRebarButton = new System.Windows.Forms.Button();
            this.utilityLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rtwRebarButton = new System.Windows.Forms.Button();
            this.creationLabel = new System.Windows.Forms.Label();
            this.TestButton = new System.Windows.Forms.Button();
            this.utilityPanel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // SetExcelButton
            // 
            this.SetExcelButton.Location = new System.Drawing.Point(595, 338);
            this.SetExcelButton.Name = "SetExcelButton";
            this.SetExcelButton.Size = new System.Drawing.Size(75, 23);
            this.SetExcelButton.TabIndex = 0;
            this.SetExcelButton.Text = "SetExcel";
            this.SetExcelButton.UseVisualStyleBackColor = true;
            this.SetExcelButton.Click += new System.EventHandler(this.OnSetExcelClick);
            // 
            // FootingRebarButton
            // 
            this.FootingRebarButton.Location = new System.Drawing.Point(25, 3);
            this.FootingRebarButton.Name = "FootingRebarButton";
            this.FootingRebarButton.Size = new System.Drawing.Size(153, 23);
            this.FootingRebarButton.TabIndex = 1;
            this.FootingRebarButton.Text = "FootingRebar";
            this.FootingRebarButton.UseVisualStyleBackColor = true;
            this.FootingRebarButton.Click += new System.EventHandler(this.OnFootingRebarClick);
            // 
            // ExcelNameLabel
            // 
            this.ExcelNameLabel.AutoSize = true;
            this.ExcelNameLabel.Location = new System.Drawing.Point(708, 343);
            this.ExcelNameLabel.Name = "ExcelNameLabel";
            this.ExcelNameLabel.Size = new System.Drawing.Size(80, 13);
            this.ExcelNameLabel.TabIndex = 2;
            this.ExcelNameLabel.Text = "No file selected";
            // 
            // SheetComboBox
            // 
            this.SheetComboBox.FormattingEnabled = true;
            this.SheetComboBox.Location = new System.Drawing.Point(595, 367);
            this.SheetComboBox.Name = "SheetComboBox";
            this.SheetComboBox.Size = new System.Drawing.Size(121, 21);
            this.SheetComboBox.TabIndex = 5;
            this.SheetComboBox.SelectedIndexChanged += new System.EventHandler(this.OnSelectedSheetChanged);
            // 
            // LoadButton
            // 
            this.LoadButton.Location = new System.Drawing.Point(595, 394);
            this.LoadButton.Name = "LoadButton";
            this.LoadButton.Size = new System.Drawing.Size(75, 23);
            this.LoadButton.TabIndex = 6;
            this.LoadButton.Text = "Load";
            this.LoadButton.UseVisualStyleBackColor = true;
            this.LoadButton.Click += new System.EventHandler(this.OnLoadButtonClick);
            // 
            // CopySpacingButton
            // 
            this.CopySpacingButton.Location = new System.Drawing.Point(52, 3);
            this.CopySpacingButton.Name = "CopySpacingButton";
            this.CopySpacingButton.Size = new System.Drawing.Size(104, 23);
            this.CopySpacingButton.TabIndex = 7;
            this.CopySpacingButton.Text = "Copy spacing";
            this.CopySpacingButton.UseVisualStyleBackColor = true;
            this.CopySpacingButton.Click += new System.EventHandler(this.OnCopySpacingClick);
            // 
            // utilityPanel
            // 
            this.utilityPanel.Controls.Add(this.recreateRebarButton);
            this.utilityPanel.Controls.Add(this.CopySpacingButton);
            this.utilityPanel.Location = new System.Drawing.Point(560, 48);
            this.utilityPanel.Name = "utilityPanel";
            this.utilityPanel.Size = new System.Drawing.Size(200, 100);
            this.utilityPanel.TabIndex = 8;
            // 
            // recreateRebarButton
            // 
            this.recreateRebarButton.Location = new System.Drawing.Point(52, 32);
            this.recreateRebarButton.Name = "recreateRebarButton";
            this.recreateRebarButton.Size = new System.Drawing.Size(104, 23);
            this.recreateRebarButton.TabIndex = 8;
            this.recreateRebarButton.Text = "Recreate Rebar";
            this.recreateRebarButton.UseVisualStyleBackColor = true;
            this.recreateRebarButton.Click += new System.EventHandler(this.OnRecreateRebarButtonClick);
            // 
            // utilityLabel
            // 
            this.utilityLabel.AutoSize = true;
            this.utilityLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.utilityLabel.Location = new System.Drawing.Point(637, 25);
            this.utilityLabel.Name = "utilityLabel";
            this.utilityLabel.Size = new System.Drawing.Size(47, 20);
            this.utilityLabel.TabIndex = 9;
            this.utilityLabel.Text = "Utility";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.rtwRebarButton);
            this.panel1.Controls.Add(this.FootingRebarButton);
            this.panel1.Location = new System.Drawing.Point(12, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 369);
            this.panel1.TabIndex = 9;
            // 
            // rtwRebarButton
            // 
            this.rtwRebarButton.Location = new System.Drawing.Point(25, 32);
            this.rtwRebarButton.Name = "rtwRebarButton";
            this.rtwRebarButton.Size = new System.Drawing.Size(153, 23);
            this.rtwRebarButton.TabIndex = 2;
            this.rtwRebarButton.Text = "Wall Rebar";
            this.rtwRebarButton.UseVisualStyleBackColor = true;
            this.rtwRebarButton.Click += new System.EventHandler(this.OnRTWButtonClick);
            // 
            // creationLabel
            // 
            this.creationLabel.AutoSize = true;
            this.creationLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.creationLabel.Location = new System.Drawing.Point(77, 25);
            this.creationLabel.Name = "creationLabel";
            this.creationLabel.Size = new System.Drawing.Size(69, 20);
            this.creationLabel.TabIndex = 10;
            this.creationLabel.Text = "Creation";
            // 
            // TestButton
            // 
            this.TestButton.Location = new System.Drawing.Point(372, 208);
            this.TestButton.Name = "TestButton";
            this.TestButton.Size = new System.Drawing.Size(75, 23);
            this.TestButton.TabIndex = 11;
            this.TestButton.Text = "Test";
            this.TestButton.UseVisualStyleBackColor = true;
            this.TestButton.Click += new System.EventHandler(this.OnTestButtonClick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.TestButton);
            this.Controls.Add(this.creationLabel);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.utilityLabel);
            this.Controls.Add(this.utilityPanel);
            this.Controls.Add(this.LoadButton);
            this.Controls.Add(this.SheetComboBox);
            this.Controls.Add(this.ExcelNameLabel);
            this.Controls.Add(this.SetExcelButton);
            this.Name = "MainForm";
            this.Text = "ZeroTouchBridge";
            this.utilityPanel.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SetExcelButton;
        private System.Windows.Forms.Button FootingRebarButton;
        public System.Windows.Forms.Label ExcelNameLabel;
        private System.Windows.Forms.ComboBox SheetComboBox;
        private System.Windows.Forms.Button LoadButton;
        private System.Windows.Forms.Button CopySpacingButton;
        private System.Windows.Forms.Panel utilityPanel;
        private System.Windows.Forms.Button recreateRebarButton;
        private System.Windows.Forms.Label utilityLabel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label creationLabel;
        private System.Windows.Forms.Button rtwRebarButton;
        private System.Windows.Forms.Button TestButton;
    }
}