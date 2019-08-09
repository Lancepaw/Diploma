namespace DataBaseIO
{
    partial class frmEditDuty
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
            this.lblDutyList = new System.Windows.Forms.Label();
            this.cmbDutyList = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDutyFromList = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.lblShort = new System.Windows.Forms.Label();
            this.txtShort = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(216, 172);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblDutyList
            // 
            this.lblDutyList.AutoSize = true;
            this.lblDutyList.Location = new System.Drawing.Point(12, 9);
            this.lblDutyList.Name = "lblDutyList";
            this.lblDutyList.Size = new System.Drawing.Size(111, 13);
            this.lblDutyList.TabIndex = 1;
            this.lblDutyList.Text = "Список Должностей";
            // 
            // cmbDutyList
            // 
            this.cmbDutyList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDutyList.FormattingEnabled = true;
            this.cmbDutyList.Location = new System.Drawing.Point(15, 25);
            this.cmbDutyList.Name = "cmbDutyList";
            this.cmbDutyList.Size = new System.Drawing.Size(276, 21);
            this.cmbDutyList.TabIndex = 2;
            this.cmbDutyList.SelectedIndexChanged += new System.EventHandler(this.cmbDutyList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Должность";
            // 
            // txtDutyFromList
            // 
            this.txtDutyFromList.Location = new System.Drawing.Point(15, 74);
            this.txtDutyFromList.Name = "txtDutyFromList";
            this.txtDutyFromList.Size = new System.Drawing.Size(276, 20);
            this.txtDutyFromList.TabIndex = 4;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(135, 143);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(216, 143);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(135, 172);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 7;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // lblShort
            // 
            this.lblShort.AutoSize = true;
            this.lblShort.Location = new System.Drawing.Point(12, 99);
            this.lblShort.Name = "lblShort";
            this.lblShort.Size = new System.Drawing.Size(132, 13);
            this.lblShort.TabIndex = 8;
            this.lblShort.Text = "Короткое наименование";
            // 
            // txtShort
            // 
            this.txtShort.Location = new System.Drawing.Point(15, 117);
            this.txtShort.Name = "txtShort";
            this.txtShort.Size = new System.Drawing.Size(276, 20);
            this.txtShort.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 177);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Всё готово";
            // 
            // frmEditDuty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 207);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtShort);
            this.Controls.Add(this.lblShort);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtDutyFromList);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbDutyList);
            this.Controls.Add(this.lblDutyList);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditDuty";
            this.Text = "Редактор должностей";
            this.Load += new System.EventHandler(this.frmEditDuty_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblDutyList;
        private System.Windows.Forms.ComboBox cmbDutyList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDutyFromList;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label lblShort;
        private System.Windows.Forms.TextBox txtShort;
        private System.Windows.Forms.Label label2;
    }
}