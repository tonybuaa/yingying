namespace yingying
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnImport = new System.Windows.Forms.Button();
            this.btnBusinessSum = new System.Windows.Forms.Button();
            this.txtSec1 = new System.Windows.Forms.TextBox();
            this.cbMonth = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnExportToWord = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.cbSourceYear = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbSourceMonth = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cbYear = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lblCurrentFile = new System.Windows.Forms.Label();
            this.btnNAS = new System.Windows.Forms.Button();
            this.dgvGenerateCaseTable = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGenerateCaseTable)).BeginInit();
            this.SuspendLayout();
            // 
            // btnImport
            // 
            this.btnImport.AutoSize = true;
            this.btnImport.Location = new System.Drawing.Point(340, 10);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(93, 23);
            this.btnImport.TabIndex = 0;
            this.btnImport.Text = "导入Excel文件";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnBusinessSum
            // 
            this.btnBusinessSum.AutoSize = true;
            this.btnBusinessSum.Enabled = false;
            this.btnBusinessSum.Location = new System.Drawing.Point(14, 110);
            this.btnBusinessSum.Name = "btnBusinessSum";
            this.btnBusinessSum.Size = new System.Drawing.Size(75, 23);
            this.btnBusinessSum.TabIndex = 1;
            this.btnBusinessSum.Text = "各行业汇总";
            this.btnBusinessSum.UseVisualStyleBackColor = true;
            this.btnBusinessSum.Click += new System.EventHandler(this.btnBusinessSum_Click);
            // 
            // txtSec1
            // 
            this.txtSec1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSec1.Location = new System.Drawing.Point(12, 139);
            this.txtSec1.Multiline = true;
            this.txtSec1.Name = "txtSec1";
            this.txtSec1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtSec1.Size = new System.Drawing.Size(771, 114);
            this.txtSec1.TabIndex = 2;
            // 
            // cbMonth
            // 
            this.cbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMonth.FormattingEnabled = true;
            this.cbMonth.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});
            this.cbMonth.Location = new System.Drawing.Point(158, 56);
            this.cbMonth.Name = "cbMonth";
            this.cbMonth.Size = new System.Drawing.Size(48, 20);
            this.cbMonth.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "报表时间：";
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(12, 571);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(771, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // btnExportToWord
            // 
            this.btnExportToWord.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportToWord.AutoSize = true;
            this.btnExportToWord.Location = new System.Drawing.Point(684, 542);
            this.btnExportToWord.Name = "btnExportToWord";
            this.btnExportToWord.Size = new System.Drawing.Size(99, 23);
            this.btnExportToWord.TabIndex = 6;
            this.btnExportToWord.Text = "Export to Word";
            this.btnExportToWord.UseVisualStyleBackColor = true;
            this.btnExportToWord.Click += new System.EventHandler(this.btnExportToWord_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "各区县Excel表时间：";
            // 
            // cbSourceYear
            // 
            this.cbSourceYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSourceYear.FormattingEnabled = true;
            this.cbSourceYear.Items.AddRange(new object[] {
            "2015",
            "2016",
            "2017",
            "2018"});
            this.cbSourceYear.Location = new System.Drawing.Point(137, 12);
            this.cbSourceYear.Name = "cbSourceYear";
            this.cbSourceYear.Size = new System.Drawing.Size(74, 20);
            this.cbSourceYear.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(217, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "年";
            // 
            // cbSourceMonth
            // 
            this.cbSourceMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSourceMonth.FormattingEnabled = true;
            this.cbSourceMonth.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});
            this.cbSourceMonth.Location = new System.Drawing.Point(240, 12);
            this.cbSourceMonth.Name = "cbSourceMonth";
            this.cbSourceMonth.Size = new System.Drawing.Size(50, 20);
            this.cbSourceMonth.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(296, 15);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(17, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "月";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(212, 62);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(17, 12);
            this.label5.TabIndex = 12;
            this.label5.Text = "月";
            // 
            // cbYear
            // 
            this.cbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbYear.FormattingEnabled = true;
            this.cbYear.Items.AddRange(new object[] {
            "2015",
            "2016",
            "2017",
            "2018"});
            this.cbYear.Location = new System.Drawing.Point(69, 56);
            this.cbYear.Name = "cbYear";
            this.cbYear.Size = new System.Drawing.Size(62, 20);
            this.cbYear.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(135, 62);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "年";
            // 
            // lblCurrentFile
            // 
            this.lblCurrentFile.AutoSize = true;
            this.lblCurrentFile.Location = new System.Drawing.Point(12, 486);
            this.lblCurrentFile.Name = "lblCurrentFile";
            this.lblCurrentFile.Size = new System.Drawing.Size(0, 12);
            this.lblCurrentFile.TabIndex = 15;
            // 
            // btnNAS
            // 
            this.btnNAS.Location = new System.Drawing.Point(95, 110);
            this.btnNAS.Name = "btnNAS";
            this.btnNAS.Size = new System.Drawing.Size(94, 23);
            this.btnNAS.TabIndex = 16;
            this.btnNAS.Text = "Connect NAS";
            this.btnNAS.UseVisualStyleBackColor = true;
            this.btnNAS.Click += new System.EventHandler(this.btnNAS_Click);
            // 
            // dgvGenerateCaseTable
            // 
            this.dgvGenerateCaseTable.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvGenerateCaseTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGenerateCaseTable.Location = new System.Drawing.Point(12, 259);
            this.dgvGenerateCaseTable.Name = "dgvGenerateCaseTable";
            this.dgvGenerateCaseTable.RowTemplate.Height = 23;
            this.dgvGenerateCaseTable.Size = new System.Drawing.Size(771, 191);
            this.dgvGenerateCaseTable.TabIndex = 17;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(795, 606);
            this.Controls.Add(this.dgvGenerateCaseTable);
            this.Controls.Add(this.btnNAS);
            this.Controls.Add(this.lblCurrentFile);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cbYear);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cbSourceMonth);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbSourceYear);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnExportToWord);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbMonth);
            this.Controls.Add(this.txtSec1);
            this.Controls.Add(this.btnBusinessSum);
            this.Controls.Add(this.btnImport);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvGenerateCaseTable)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnBusinessSum;
        private System.Windows.Forms.TextBox txtSec1;
        private System.Windows.Forms.ComboBox cbMonth;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnExportToWord;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbSourceYear;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbSourceMonth;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbYear;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblCurrentFile;
        private System.Windows.Forms.Button btnNAS;
        private System.Windows.Forms.DataGridView dgvGenerateCaseTable;
    }
}

