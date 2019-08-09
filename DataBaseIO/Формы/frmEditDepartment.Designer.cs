namespace DataBaseIO
{
    partial class frmEditDepartment
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
            this.cmbDepartList = new System.Windows.Forms.ComboBox();
            this.lblDepartList = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtShort = new System.Windows.Forms.TextBox();
            this.lblShort = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(312, 159);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // cmbDepartList
            // 
            this.cmbDepartList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDepartList.FormattingEnabled = true;
            this.cmbDepartList.Location = new System.Drawing.Point(12, 25);
            this.cmbDepartList.Name = "cmbDepartList";
            this.cmbDepartList.Size = new System.Drawing.Size(375, 21);
            this.cmbDepartList.TabIndex = 1;
            this.cmbDepartList.SelectedIndexChanged += new System.EventHandler(this.cmbDepartList_SelectedIndexChanged);
            // 
            // lblDepartList
            // 
            this.lblDepartList.AutoSize = true;
            this.lblDepartList.Location = new System.Drawing.Point(9, 9);
            this.lblDepartList.Name = "lblDepartList";
            this.lblDepartList.Size = new System.Drawing.Size(85, 13);
            this.lblDepartList.TabIndex = 2;
            this.lblDepartList.Text = "Список кафедр";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(12, 65);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(375, 20);
            this.txtName.TabIndex = 3;
            // 
            // txtShort
            // 
            this.txtShort.Location = new System.Drawing.Point(12, 104);
            this.txtShort.Name = "txtShort";
            this.txtShort.Size = new System.Drawing.Size(375, 20);
            this.txtShort.TabIndex = 4;
            // 
            // lblShort
            // 
            this.lblShort.AutoSize = true;
            this.lblShort.Location = new System.Drawing.Point(12, 88);
            this.lblShort.Name = "lblShort";
            this.lblShort.Size = new System.Drawing.Size(106, 13);
            this.lblShort.TabIndex = 5;
            this.lblShort.Text = "Короткое название";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(12, 49);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(96, 13);
            this.lblName.TabIndex = 6;
            this.lblName.Text = "Полное название";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(312, 130);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(231, 130);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 8;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(231, 159);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 9;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 164);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Всё готово";
            // 
            // frmEditDepartment
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(395, 190);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.lblShort);
            this.Controls.Add(this.txtShort);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.lblDepartList);
            this.Controls.Add(this.cmbDepartList);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditDepartment";
            this.Text = "Редактор кафедр";
            this.Load += new System.EventHandler(this.frmEditDepartment_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ComboBox cmbDepartList;
        private System.Windows.Forms.Label lblDepartList;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtShort;
        private System.Windows.Forms.Label lblShort;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label label1;
    }
}