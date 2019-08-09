namespace DataBaseIO
{
    partial class frmEditWorkYear
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
            this.btnClose = new System.Windows.Forms.Button();
            this.lblWorkYearList = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.cmbWorkYearList = new System.Windows.Forms.ComboBox();
            this.txtWorkYear = new System.Windows.Forms.TextBox();
            this.lblWorkYear = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(96, 120);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblWorkYearList
            // 
            this.lblWorkYearList.AutoSize = true;
            this.lblWorkYearList.Location = new System.Drawing.Point(12, 9);
            this.lblWorkYearList.Name = "lblWorkYearList";
            this.lblWorkYearList.Size = new System.Drawing.Size(72, 13);
            this.lblWorkYearList.TabIndex = 1;
            this.lblWorkYearList.Text = "Учебный год";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(96, 91);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(15, 91);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(15, 120);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 4;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // cmbWorkYearList
            // 
            this.cmbWorkYearList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWorkYearList.FormattingEnabled = true;
            this.cmbWorkYearList.Location = new System.Drawing.Point(15, 25);
            this.cmbWorkYearList.Name = "cmbWorkYearList";
            this.cmbWorkYearList.Size = new System.Drawing.Size(156, 21);
            this.cmbWorkYearList.TabIndex = 5;
            this.cmbWorkYearList.SelectedIndexChanged += new System.EventHandler(this.cmbWorkYearList_SelectedIndexChanged);
            // 
            // txtWorkYear
            // 
            this.txtWorkYear.Location = new System.Drawing.Point(15, 65);
            this.txtWorkYear.Name = "txtWorkYear";
            this.txtWorkYear.Size = new System.Drawing.Size(156, 20);
            this.txtWorkYear.TabIndex = 6;
            // 
            // lblWorkYear
            // 
            this.lblWorkYear.AutoSize = true;
            this.lblWorkYear.Location = new System.Drawing.Point(12, 49);
            this.lblWorkYear.Name = "lblWorkYear";
            this.lblWorkYear.Size = new System.Drawing.Size(72, 13);
            this.lblWorkYear.TabIndex = 7;
            this.lblWorkYear.Text = "Учебный год";
            // 
            // frmEditWorkYear
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(183, 158);
            this.Controls.Add(this.lblWorkYear);
            this.Controls.Add(this.txtWorkYear);
            this.Controls.Add(this.cmbWorkYearList);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblWorkYearList);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditWorkYear";
            this.Text = "Редактор учебных годов";
            this.Load += new System.EventHandler(this.frmEditWorkYear_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblWorkYearList;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.ComboBox cmbWorkYearList;
        private System.Windows.Forms.TextBox txtWorkYear;
        private System.Windows.Forms.Label lblWorkYear;
    }
}