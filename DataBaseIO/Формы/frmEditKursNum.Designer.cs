namespace DataBaseIO
{
    partial class frmEditKursNum
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
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.cmbKursNumList = new System.Windows.Forms.ComboBox();
            this.lblKursNumList = new System.Windows.Forms.Label();
            this.lblKursNum = new System.Windows.Forms.Label();
            this.txtKursNum = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(187, 146);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(187, 116);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(106, 116);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(106, 146);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // cmbKursNumList
            // 
            this.cmbKursNumList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbKursNumList.FormattingEnabled = true;
            this.cmbKursNumList.Location = new System.Drawing.Point(12, 25);
            this.cmbKursNumList.Name = "cmbKursNumList";
            this.cmbKursNumList.Size = new System.Drawing.Size(250, 21);
            this.cmbKursNumList.TabIndex = 4;
            this.cmbKursNumList.SelectedIndexChanged += new System.EventHandler(this.cmbKursNumList_SelectedIndexChanged);
            // 
            // lblKursNumList
            // 
            this.lblKursNumList.AutoSize = true;
            this.lblKursNumList.Location = new System.Drawing.Point(12, 9);
            this.lblKursNumList.Name = "lblKursNumList";
            this.lblKursNumList.Size = new System.Drawing.Size(73, 13);
            this.lblKursNumList.TabIndex = 5;
            this.lblKursNumList.Text = "Номер курса";
            // 
            // lblKursNum
            // 
            this.lblKursNum.AutoSize = true;
            this.lblKursNum.Location = new System.Drawing.Point(12, 49);
            this.lblKursNum.Name = "lblKursNum";
            this.lblKursNum.Size = new System.Drawing.Size(73, 13);
            this.lblKursNum.TabIndex = 6;
            this.lblKursNum.Text = "Номер курса";
            // 
            // txtKursNum
            // 
            this.txtKursNum.Location = new System.Drawing.Point(12, 65);
            this.txtKursNum.Name = "txtKursNum";
            this.txtKursNum.Size = new System.Drawing.Size(250, 20);
            this.txtKursNum.TabIndex = 7;
            // 
            // frmEditKursNum
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(274, 181);
            this.Controls.Add(this.txtKursNum);
            this.Controls.Add(this.lblKursNum);
            this.Controls.Add(this.lblKursNumList);
            this.Controls.Add(this.cmbKursNumList);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditKursNum";
            this.Text = "Редактор номеров курсов";
            this.Load += new System.EventHandler(this.frmEditKursNum_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.ComboBox cmbKursNumList;
        private System.Windows.Forms.Label lblKursNumList;
        private System.Windows.Forms.Label lblKursNum;
        private System.Windows.Forms.TextBox txtKursNum;
    }
}