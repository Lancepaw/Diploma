namespace DataBaseIO
{
    partial class frmEditSemestr
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
            this.lblSemestrList = new System.Windows.Forms.Label();
            this.cmbSemestrList = new System.Windows.Forms.ComboBox();
            this.txtSemNum = new System.Windows.Forms.TextBox();
            this.lblSemNum = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(93, 121);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblSemestrList
            // 
            this.lblSemestrList.AutoSize = true;
            this.lblSemestrList.Location = new System.Drawing.Point(12, 9);
            this.lblSemestrList.Name = "lblSemestrList";
            this.lblSemestrList.Size = new System.Drawing.Size(51, 13);
            this.lblSemestrList.TabIndex = 1;
            this.lblSemestrList.Text = "Семестр";
            // 
            // cmbSemestrList
            // 
            this.cmbSemestrList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestrList.FormattingEnabled = true;
            this.cmbSemestrList.Location = new System.Drawing.Point(12, 25);
            this.cmbSemestrList.Name = "cmbSemestrList";
            this.cmbSemestrList.Size = new System.Drawing.Size(156, 21);
            this.cmbSemestrList.TabIndex = 2;
            this.cmbSemestrList.SelectedIndexChanged += new System.EventHandler(this.cmbSemestrList_SelectedIndexChanged);
            // 
            // txtSemNum
            // 
            this.txtSemNum.Location = new System.Drawing.Point(12, 65);
            this.txtSemNum.Name = "txtSemNum";
            this.txtSemNum.Size = new System.Drawing.Size(156, 20);
            this.txtSemNum.TabIndex = 3;
            // 
            // lblSemNum
            // 
            this.lblSemNum.AutoSize = true;
            this.lblSemNum.Location = new System.Drawing.Point(12, 49);
            this.lblSemNum.Name = "lblSemNum";
            this.lblSemNum.Size = new System.Drawing.Size(93, 13);
            this.lblSemNum.TabIndex = 4;
            this.lblSemNum.Text = "Номер семестра";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(93, 91);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 24);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(12, 91);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(12, 120);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 7;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(90, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Всё готово";
            // 
            // frmEditSemestr
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(177, 153);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblSemNum);
            this.Controls.Add(this.txtSemNum);
            this.Controls.Add(this.cmbSemestrList);
            this.Controls.Add(this.lblSemestrList);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditSemestr";
            this.Text = "Редактор семестров";
            this.Load += new System.EventHandler(this.frmEditSemestr_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblSemestrList;
        private System.Windows.Forms.ComboBox cmbSemestrList;
        private System.Windows.Forms.TextBox txtSemNum;
        private System.Windows.Forms.Label lblSemNum;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label label1;
    }
}