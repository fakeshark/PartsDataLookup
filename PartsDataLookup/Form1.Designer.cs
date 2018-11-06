namespace PartsDataLookup
{
    partial class Form1
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
            this.btnLoadPartsSearchList = new System.Windows.Forms.Button();
            this.btnLoadFlisFoi = new System.Windows.Forms.Button();
            this.dgvExcelList = new System.Windows.Forms.DataGridView();
            this.dgvMatchList = new System.Windows.Forms.DataGridView();
            this.lblPartsToMatch = new System.Windows.Forms.Label();
            this.lblMatchingRecords = new System.Windows.Forms.Label();
            this.btnMatchRecords = new System.Windows.Forms.Button();
            this.lblMatchResults = new System.Windows.Forms.Label();
            this.btnExportToExcell = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnLoad036 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMatchList)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLoadPartsSearchList
            // 
            this.btnLoadPartsSearchList.Location = new System.Drawing.Point(12, 12);
            this.btnLoadPartsSearchList.Name = "btnLoadPartsSearchList";
            this.btnLoadPartsSearchList.Size = new System.Drawing.Size(279, 40);
            this.btnLoadPartsSearchList.TabIndex = 0;
            this.btnLoadPartsSearchList.Text = "Load Parts Search Term List";
            this.btnLoadPartsSearchList.UseVisualStyleBackColor = true;
            this.btnLoadPartsSearchList.Click += new System.EventHandler(this.LoadPartsSearchList_Click);
            // 
            // btnLoadFlisFoi
            // 
            this.btnLoadFlisFoi.Enabled = false;
            this.btnLoadFlisFoi.Location = new System.Drawing.Point(12, 87);
            this.btnLoadFlisFoi.Name = "btnLoadFlisFoi";
            this.btnLoadFlisFoi.Size = new System.Drawing.Size(279, 40);
            this.btnLoadFlisFoi.TabIndex = 1;
            this.btnLoadFlisFoi.Text = "Update FLIS Packaging Data File";
            this.btnLoadFlisFoi.UseVisualStyleBackColor = true;
            this.btnLoadFlisFoi.Click += new System.EventHandler(this.LoadFlisFoi_Click);
            // 
            // dgvExcelList
            // 
            this.dgvExcelList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvExcelList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvExcelList.Location = new System.Drawing.Point(297, 13);
            this.dgvExcelList.Name = "dgvExcelList";
            this.dgvExcelList.Size = new System.Drawing.Size(924, 304);
            this.dgvExcelList.TabIndex = 4;
            // 
            // dgvMatchList
            // 
            this.dgvMatchList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvMatchList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMatchList.Location = new System.Drawing.Point(12, 323);
            this.dgvMatchList.Name = "dgvMatchList";
            this.dgvMatchList.Size = new System.Drawing.Size(1209, 242);
            this.dgvMatchList.TabIndex = 5;
            // 
            // lblPartsToMatch
            // 
            this.lblPartsToMatch.AutoSize = true;
            this.lblPartsToMatch.Location = new System.Drawing.Point(12, 55);
            this.lblPartsToMatch.Name = "lblPartsToMatch";
            this.lblPartsToMatch.Size = new System.Drawing.Size(16, 13);
            this.lblPartsToMatch.TabIndex = 6;
            this.lblPartsToMatch.Text = "...";
            // 
            // lblMatchingRecords
            // 
            this.lblMatchingRecords.AutoSize = true;
            this.lblMatchingRecords.Location = new System.Drawing.Point(12, 130);
            this.lblMatchingRecords.Name = "lblMatchingRecords";
            this.lblMatchingRecords.Size = new System.Drawing.Size(16, 13);
            this.lblMatchingRecords.TabIndex = 7;
            this.lblMatchingRecords.Text = "...";
            // 
            // btnMatchRecords
            // 
            this.btnMatchRecords.Enabled = false;
            this.btnMatchRecords.Location = new System.Drawing.Point(12, 162);
            this.btnMatchRecords.Name = "btnMatchRecords";
            this.btnMatchRecords.Size = new System.Drawing.Size(279, 40);
            this.btnMatchRecords.TabIndex = 8;
            this.btnMatchRecords.Text = "Match Part Numbers";
            this.btnMatchRecords.UseVisualStyleBackColor = true;
            this.btnMatchRecords.Click += new System.EventHandler(this.MatchRecords_Click);
            // 
            // lblMatchResults
            // 
            this.lblMatchResults.AutoSize = true;
            this.lblMatchResults.Location = new System.Drawing.Point(12, 205);
            this.lblMatchResults.Name = "lblMatchResults";
            this.lblMatchResults.Size = new System.Drawing.Size(16, 13);
            this.lblMatchResults.TabIndex = 9;
            this.lblMatchResults.Text = "...";
            // 
            // btnExportToExcell
            // 
            this.btnExportToExcell.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnExportToExcell.Enabled = false;
            this.btnExportToExcell.Location = new System.Drawing.Point(477, 578);
            this.btnExportToExcell.Name = "btnExportToExcell";
            this.btnExportToExcell.Size = new System.Drawing.Size(279, 40);
            this.btnExportToExcell.TabIndex = 10;
            this.btnExportToExcell.Text = "Export To Spreadsheet (.xlsx)";
            this.btnExportToExcell.UseVisualStyleBackColor = true;
            this.btnExportToExcell.Click += new System.EventHandler(this.ExportToExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 280);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "...";
            // 
            // btnLoad036
            // 
            this.btnLoad036.Location = new System.Drawing.Point(12, 237);
            this.btnLoad036.Name = "btnLoad036";
            this.btnLoad036.Size = new System.Drawing.Size(279, 40);
            this.btnLoad036.TabIndex = 11;
            this.btnLoad036.Text = "Load Additional Parts Data Files (.036)";
            this.btnLoad036.UseVisualStyleBackColor = true;
            this.btnLoad036.Click += new System.EventHandler(this.Load036_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1233, 630);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLoad036);
            this.Controls.Add(this.btnExportToExcell);
            this.Controls.Add(this.lblMatchResults);
            this.Controls.Add(this.btnMatchRecords);
            this.Controls.Add(this.lblMatchingRecords);
            this.Controls.Add(this.lblPartsToMatch);
            this.Controls.Add(this.dgvMatchList);
            this.Controls.Add(this.dgvExcelList);
            this.Controls.Add(this.btnLoadFlisFoi);
            this.Controls.Add(this.btnLoadPartsSearchList);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvExcelList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMatchList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoadPartsSearchList;
        private System.Windows.Forms.Button btnLoadFlisFoi;
        public System.Windows.Forms.DataGridView dgvExcelList;
        public System.Windows.Forms.DataGridView dgvMatchList;
        private System.Windows.Forms.Label lblPartsToMatch;
        private System.Windows.Forms.Label lblMatchingRecords;
        private System.Windows.Forms.Button btnMatchRecords;
        private System.Windows.Forms.Label lblMatchResults;
        private System.Windows.Forms.Button btnExportToExcell;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnLoad036;
    }
}

