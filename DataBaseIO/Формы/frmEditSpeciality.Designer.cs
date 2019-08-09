namespace DataBaseIO
{
    partial class frmEditSpeciality
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
            this.btnDel = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lblSpecialityList = new System.Windows.Forms.Label();
            this.cmbSpecialityList = new System.Windows.Forms.ComboBox();
            this.lblSpeciality = new System.Windows.Forms.Label();
            this.txtSpecialityShort = new System.Windows.Forms.TextBox();
            this.txtSpecialityFull = new System.Windows.Forms.TextBox();
            this.lblSpecShort = new System.Windows.Forms.Label();
            this.txtSpecShort = new System.Windows.Forms.TextBox();
            this.lblInst = new System.Windows.Forms.Label();
            this.txtInst = new System.Windows.Forms.TextBox();
            this.lblFaculty = new System.Windows.Forms.Label();
            this.cmbFaculty = new System.Windows.Forms.ComboBox();
            this.lblSpecialityShort = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDiff = new System.Windows.Forms.Label();
            this.txtDiff = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(278, 206);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(278, 177);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(197, 206);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 2;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(197, 177);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lblSpecialityList
            // 
            this.lblSpecialityList.AutoSize = true;
            this.lblSpecialityList.Location = new System.Drawing.Point(12, 9);
            this.lblSpecialityList.Name = "lblSpecialityList";
            this.lblSpecialityList.Size = new System.Drawing.Size(136, 13);
            this.lblSpecialityList.TabIndex = 4;
            this.lblSpecialityList.Text = "Специализация (коротко)";
            // 
            // cmbSpecialityList
            // 
            this.cmbSpecialityList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpecialityList.FormattingEnabled = true;
            this.cmbSpecialityList.Location = new System.Drawing.Point(12, 25);
            this.cmbSpecialityList.Name = "cmbSpecialityList";
            this.cmbSpecialityList.Size = new System.Drawing.Size(135, 21);
            this.cmbSpecialityList.TabIndex = 5;
            this.cmbSpecialityList.SelectedIndexChanged += new System.EventHandler(this.cmbSpecialityList_SelectedIndexChanged);
            // 
            // lblSpeciality
            // 
            this.lblSpeciality.AutoSize = true;
            this.lblSpeciality.Location = new System.Drawing.Point(12, 49);
            this.lblSpeciality.Name = "lblSpeciality";
            this.lblSpeciality.Size = new System.Drawing.Size(202, 13);
            this.lblSpeciality.TabIndex = 6;
            this.lblSpeciality.Text = "Полное наименование специальности";
            // 
            // txtSpecialityShort
            // 
            this.txtSpecialityShort.Location = new System.Drawing.Point(153, 26);
            this.txtSpecialityShort.Name = "txtSpecialityShort";
            this.txtSpecialityShort.Size = new System.Drawing.Size(200, 20);
            this.txtSpecialityShort.TabIndex = 7;
            // 
            // txtSpecialityFull
            // 
            this.txtSpecialityFull.Location = new System.Drawing.Point(12, 65);
            this.txtSpecialityFull.Name = "txtSpecialityFull";
            this.txtSpecialityFull.Size = new System.Drawing.Size(341, 20);
            this.txtSpecialityFull.TabIndex = 8;
            // 
            // lblSpecShort
            // 
            this.lblSpecShort.AutoSize = true;
            this.lblSpecShort.Location = new System.Drawing.Point(12, 88);
            this.lblSpecShort.Name = "lblSpecShort";
            this.lblSpecShort.Size = new System.Drawing.Size(135, 13);
            this.lblSpecShort.TabIndex = 9;
            this.lblSpecShort.Text = "Специальность (коротко)";
            // 
            // txtSpecShort
            // 
            this.txtSpecShort.Location = new System.Drawing.Point(12, 104);
            this.txtSpecShort.Name = "txtSpecShort";
            this.txtSpecShort.Size = new System.Drawing.Size(135, 20);
            this.txtSpecShort.TabIndex = 10;
            // 
            // lblInst
            // 
            this.lblInst.AutoSize = true;
            this.lblInst.Location = new System.Drawing.Point(153, 88);
            this.lblInst.Name = "lblInst";
            this.lblInst.Size = new System.Drawing.Size(150, 13);
            this.lblInst.TabIndex = 11;
            this.lblInst.Text = "Институтская аббревиатура";
            // 
            // txtInst
            // 
            this.txtInst.Location = new System.Drawing.Point(153, 104);
            this.txtInst.Name = "txtInst";
            this.txtInst.Size = new System.Drawing.Size(200, 20);
            this.txtInst.TabIndex = 12;
            // 
            // lblFaculty
            // 
            this.lblFaculty.AutoSize = true;
            this.lblFaculty.Location = new System.Drawing.Point(12, 127);
            this.lblFaculty.Name = "lblFaculty";
            this.lblFaculty.Size = new System.Drawing.Size(63, 13);
            this.lblFaculty.TabIndex = 13;
            this.lblFaculty.Text = "Факультет";
            // 
            // cmbFaculty
            // 
            this.cmbFaculty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFaculty.FormattingEnabled = true;
            this.cmbFaculty.Location = new System.Drawing.Point(12, 143);
            this.cmbFaculty.Name = "cmbFaculty";
            this.cmbFaculty.Size = new System.Drawing.Size(341, 21);
            this.cmbFaculty.TabIndex = 14;
            // 
            // lblSpecialityShort
            // 
            this.lblSpecialityShort.AutoSize = true;
            this.lblSpecialityShort.Location = new System.Drawing.Point(154, 9);
            this.lblSpecialityShort.Name = "lblSpecialityShort";
            this.lblSpecialityShort.Size = new System.Drawing.Size(135, 13);
            this.lblSpecialityShort.TabIndex = 15;
            this.lblSpecialityShort.Text = "Специальность (коротко)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 219);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "Всё готово";
            // 
            // lblDiff
            // 
            this.lblDiff.AutoSize = true;
            this.lblDiff.Location = new System.Drawing.Point(12, 167);
            this.lblDiff.Name = "lblDiff";
            this.lblDiff.Size = new System.Drawing.Size(144, 13);
            this.lblDiff.TabIndex = 17;
            this.lblDiff.Text = "Образовательная система";
            // 
            // txtDiff
            // 
            this.txtDiff.Location = new System.Drawing.Point(12, 183);
            this.txtDiff.Name = "txtDiff";
            this.txtDiff.Size = new System.Drawing.Size(144, 20);
            this.txtDiff.TabIndex = 18;
            // 
            // frmEditSpeciality
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 241);
            this.Controls.Add(this.txtDiff);
            this.Controls.Add(this.lblDiff);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblSpecialityShort);
            this.Controls.Add(this.cmbFaculty);
            this.Controls.Add(this.lblFaculty);
            this.Controls.Add(this.txtInst);
            this.Controls.Add(this.lblInst);
            this.Controls.Add(this.txtSpecShort);
            this.Controls.Add(this.lblSpecShort);
            this.Controls.Add(this.txtSpecialityFull);
            this.Controls.Add(this.txtSpecialityShort);
            this.Controls.Add(this.lblSpeciality);
            this.Controls.Add(this.cmbSpecialityList);
            this.Controls.Add(this.lblSpecialityList);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditSpeciality";
            this.Text = "Редактор специальностей";
            this.Load += new System.EventHandler(this.frmEditSpeciality_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label lblSpecialityList;
        private System.Windows.Forms.ComboBox cmbSpecialityList;
        private System.Windows.Forms.Label lblSpeciality;
        private System.Windows.Forms.TextBox txtSpecialityShort;
        private System.Windows.Forms.TextBox txtSpecialityFull;
        private System.Windows.Forms.Label lblSpecShort;
        private System.Windows.Forms.TextBox txtSpecShort;
        private System.Windows.Forms.Label lblInst;
        private System.Windows.Forms.TextBox txtInst;
        private System.Windows.Forms.Label lblFaculty;
        private System.Windows.Forms.ComboBox cmbFaculty;
        private System.Windows.Forms.Label lblSpecialityShort;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDiff;
        private System.Windows.Forms.TextBox txtDiff;
    }
}