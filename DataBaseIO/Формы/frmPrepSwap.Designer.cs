namespace DataBaseIO
{
    partial class frmPrepSwap
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
            this.dgSwap = new System.Windows.Forms.DataGridView();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSemestr2 = new System.Windows.Forms.Button();
            this.btnSemestr1 = new System.Windows.Forms.Button();
            this.btnWord = new System.Windows.Forms.Button();
            this.lblSem = new System.Windows.Forms.Label();
            this.cmbSemestr = new System.Windows.Forms.ComboBox();
            this.lblWorkYear = new System.Windows.Forms.Label();
            this.cmbWorkYear = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgSwap)).BeginInit();
            this.SuspendLayout();
            // 
            // dgSwap
            // 
            this.dgSwap.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgSwap.Location = new System.Drawing.Point(12, 33);
            this.dgSwap.Name = "dgSwap";
            this.dgSwap.Size = new System.Drawing.Size(792, 287);
            this.dgSwap.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(729, 326);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSemestr2
            // 
            this.btnSemestr2.Location = new System.Drawing.Point(142, 326);
            this.btnSemestr2.Name = "btnSemestr2";
            this.btnSemestr2.Size = new System.Drawing.Size(124, 23);
            this.btnSemestr2.TabIndex = 2;
            this.btnSemestr2.Text = "График на II семестр";
            this.btnSemestr2.UseVisualStyleBackColor = true;
            this.btnSemestr2.Click += new System.EventHandler(this.btnSemestr2_Click);
            // 
            // btnSemestr1
            // 
            this.btnSemestr1.Location = new System.Drawing.Point(12, 326);
            this.btnSemestr1.Name = "btnSemestr1";
            this.btnSemestr1.Size = new System.Drawing.Size(124, 23);
            this.btnSemestr1.TabIndex = 3;
            this.btnSemestr1.Text = "График на I семестр";
            this.btnSemestr1.UseVisualStyleBackColor = true;
            this.btnSemestr1.Click += new System.EventHandler(this.btnSemestr1_Click);
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(272, 326);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(75, 23);
            this.btnWord.TabIndex = 4;
            this.btnWord.Text = "В Word";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // lblSem
            // 
            this.lblSem.AutoSize = true;
            this.lblSem.Location = new System.Drawing.Point(9, 9);
            this.lblSem.Name = "lblSem";
            this.lblSem.Size = new System.Drawing.Size(51, 13);
            this.lblSem.TabIndex = 5;
            this.lblSem.Text = "Семестр";
            // 
            // cmbSemestr
            // 
            this.cmbSemestr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestr.FormattingEnabled = true;
            this.cmbSemestr.Location = new System.Drawing.Point(66, 6);
            this.cmbSemestr.Name = "cmbSemestr";
            this.cmbSemestr.Size = new System.Drawing.Size(121, 21);
            this.cmbSemestr.TabIndex = 6;
            // 
            // lblWorkYear
            // 
            this.lblWorkYear.AutoSize = true;
            this.lblWorkYear.Location = new System.Drawing.Point(193, 9);
            this.lblWorkYear.Name = "lblWorkYear";
            this.lblWorkYear.Size = new System.Drawing.Size(72, 13);
            this.lblWorkYear.TabIndex = 7;
            this.lblWorkYear.Text = "Учебный год";
            // 
            // cmbWorkYear
            // 
            this.cmbWorkYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWorkYear.FormattingEnabled = true;
            this.cmbWorkYear.Location = new System.Drawing.Point(271, 6);
            this.cmbWorkYear.Name = "cmbWorkYear";
            this.cmbWorkYear.Size = new System.Drawing.Size(121, 21);
            this.cmbWorkYear.TabIndex = 8;
            // 
            // frmPrepSwap
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(816, 356);
            this.Controls.Add(this.cmbWorkYear);
            this.Controls.Add(this.lblWorkYear);
            this.Controls.Add(this.cmbSemestr);
            this.Controls.Add(this.lblSem);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.btnSemestr1);
            this.Controls.Add(this.btnSemestr2);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.dgSwap);
            this.Name = "frmPrepSwap";
            this.Text = "Замены преподавателей";
            this.Load += new System.EventHandler(this.frmPrepSwap_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgSwap)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgSwap;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSemestr2;
        private System.Windows.Forms.Button btnSemestr1;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Label lblSem;
        private System.Windows.Forms.ComboBox cmbSemestr;
        private System.Windows.Forms.Label lblWorkYear;
        private System.Windows.Forms.ComboBox cmbWorkYear;
    }
}