namespace CombinationOfExcelColumns
{
    partial class FrmCombination
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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.nudHeaderEndRow = new System.Windows.Forms.NumericUpDown();
            this.nudHeaderStartRow = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.nudWorkSheetNum = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblFileName = new System.Windows.Forms.Label();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnCreateCombinations = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.nudDataStartRow = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.ofdExcel = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudHeaderEndRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudHeaderStartRow)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudWorkSheetNum)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudDataStartRow)).BeginInit();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 360);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.statusStrip1.Size = new System.Drawing.Size(321, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.nudHeaderEndRow);
            this.groupBox1.Controls.Add(this.nudHeaderStartRow);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(7, 139);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Size = new System.Drawing.Size(304, 95);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "3-Column Header";
            // 
            // nudHeaderEndRow
            // 
            this.nudHeaderEndRow.Location = new System.Drawing.Point(112, 59);
            this.nudHeaderEndRow.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.nudHeaderEndRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudHeaderEndRow.Name = "nudHeaderEndRow";
            this.nudHeaderEndRow.Size = new System.Drawing.Size(164, 25);
            this.nudHeaderEndRow.TabIndex = 1;
            this.nudHeaderEndRow.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // nudHeaderStartRow
            // 
            this.nudHeaderStartRow.Location = new System.Drawing.Point(112, 26);
            this.nudHeaderStartRow.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.nudHeaderStartRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudHeaderStartRow.Name = "nudHeaderStartRow";
            this.nudHeaderStartRow.Size = new System.Drawing.Size(164, 25);
            this.nudHeaderStartRow.TabIndex = 1;
            this.nudHeaderStartRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "End Row";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Start Row";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.nudWorkSheetNum);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(7, 72);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox2.Size = new System.Drawing.Size(304, 59);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "2-Work Sheet";
            // 
            // nudWorkSheetNum
            // 
            this.nudWorkSheetNum.Location = new System.Drawing.Point(112, 22);
            this.nudWorkSheetNum.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.nudWorkSheetNum.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudWorkSheetNum.Name = "nudWorkSheetNum";
            this.nudWorkSheetNum.Size = new System.Drawing.Size(164, 25);
            this.nudWorkSheetNum.TabIndex = 1;
            this.nudWorkSheetNum.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 17);
            this.label4.TabIndex = 0;
            this.label4.Text = "Number";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblFileName);
            this.groupBox3.Controls.Add(this.btnSelectFile);
            this.groupBox3.Location = new System.Drawing.Point(7, 5);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox3.Size = new System.Drawing.Size(304, 59);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "1-Excel";
            // 
            // lblFileName
            // 
            this.lblFileName.Font = new System.Drawing.Font("Segoe UI", 7F);
            this.lblFileName.Location = new System.Drawing.Point(109, 21);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(189, 27);
            this.lblFileName.TabIndex = 1;
            this.lblFileName.Text = "File Path";
            this.lblFileName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(10, 21);
            this.btnSelectFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(87, 30);
            this.btnSelectFile.TabIndex = 0;
            this.btnSelectFile.Text = "Select File";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // btnCreateCombinations
            // 
            this.btnCreateCombinations.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCreateCombinations.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnCreateCombinations.ForeColor = System.Drawing.Color.Black;
            this.btnCreateCombinations.Location = new System.Drawing.Point(7, 309);
            this.btnCreateCombinations.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCreateCombinations.Name = "btnCreateCombinations";
            this.btnCreateCombinations.Size = new System.Drawing.Size(304, 46);
            this.btnCreateCombinations.TabIndex = 4;
            this.btnCreateCombinations.Text = "CREATE COMBINATIONS";
            this.btnCreateCombinations.UseVisualStyleBackColor = true;
            this.btnCreateCombinations.Click += new System.EventHandler(this.btnCreateCombinations_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.nudDataStartRow);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(7, 242);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox4.Size = new System.Drawing.Size(304, 59);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "4-Data";
            // 
            // nudDataStartRow
            // 
            this.nudDataStartRow.Location = new System.Drawing.Point(112, 20);
            this.nudDataStartRow.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.nudDataStartRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudDataStartRow.Name = "nudDataStartRow";
            this.nudDataStartRow.Size = new System.Drawing.Size(164, 25);
            this.nudDataStartRow.TabIndex = 1;
            this.nudDataStartRow.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 17);
            this.label3.TabIndex = 0;
            this.label3.Text = "Start Row";
            // 
            // ofdExcel
            // 
            this.ofdExcel.FileName = "openFileDialog1";
            // 
            // FrmCombination
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(321, 382);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.btnCreateCombinations);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "FrmCombination";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Unique Combination Of Excel Columns.";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudHeaderEndRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudHeaderStartRow)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudWorkSheetNum)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudDataStartRow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.NumericUpDown nudHeaderEndRow;
        private System.Windows.Forms.NumericUpDown nudHeaderStartRow;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.NumericUpDown nudWorkSheetNum;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnCreateCombinations;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.NumericUpDown nudDataStartRow;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.OpenFileDialog ofdExcel;
    }
}