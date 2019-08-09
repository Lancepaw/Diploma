namespace DataBaseIO
{
    partial class frmEditSickList
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
            this.lblLists = new System.Windows.Forms.Label();
            this.cmbLists = new System.Windows.Forms.ComboBox();
            this.lblLecturer = new System.Windows.Forms.Label();
            this.cmbLecturerList = new System.Windows.Forms.ComboBox();
            this.lblSemestr = new System.Windows.Forms.Label();
            this.cmbSemestrList = new System.Windows.Forms.ComboBox();
            this.dtpClose = new System.Windows.Forms.DateTimePicker();
            this.dtpOpen = new System.Windows.Forms.DateTimePicker();
            this.lblOpen = new System.Windows.Forms.Label();
            this.lblClose = new System.Windows.Forms.Label();
            this.lblDescript = new System.Windows.Forms.Label();
            this.txtDescript = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(406, 227);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(406, 198);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(325, 198);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(325, 227);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            // 
            // lblLists
            // 
            this.lblLists.AutoSize = true;
            this.lblLists.Location = new System.Drawing.Point(9, 9);
            this.lblLists.Name = "lblLists";
            this.lblLists.Size = new System.Drawing.Size(106, 13);
            this.lblLists.TabIndex = 4;
            this.lblLists.Text = "Больничные листы:";
            // 
            // cmbLists
            // 
            this.cmbLists.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLists.FormattingEnabled = true;
            this.cmbLists.Location = new System.Drawing.Point(12, 25);
            this.cmbLists.Name = "cmbLists";
            this.cmbLists.Size = new System.Drawing.Size(469, 21);
            this.cmbLists.TabIndex = 5;
            this.cmbLists.SelectedIndexChanged += new System.EventHandler(this.cmbLists_SelectedIndexChanged);
            // 
            // lblLecturer
            // 
            this.lblLecturer.AutoSize = true;
            this.lblLecturer.Location = new System.Drawing.Point(12, 49);
            this.lblLecturer.Name = "lblLecturer";
            this.lblLecturer.Size = new System.Drawing.Size(89, 13);
            this.lblLecturer.TabIndex = 6;
            this.lblLecturer.Text = "Преподаватель:";
            // 
            // cmbLecturerList
            // 
            this.cmbLecturerList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLecturerList.FormattingEnabled = true;
            this.cmbLecturerList.Location = new System.Drawing.Point(12, 65);
            this.cmbLecturerList.Name = "cmbLecturerList";
            this.cmbLecturerList.Size = new System.Drawing.Size(165, 21);
            this.cmbLecturerList.TabIndex = 7;
            // 
            // lblSemestr
            // 
            this.lblSemestr.AutoSize = true;
            this.lblSemestr.Location = new System.Drawing.Point(180, 49);
            this.lblSemestr.Name = "lblSemestr";
            this.lblSemestr.Size = new System.Drawing.Size(54, 13);
            this.lblSemestr.TabIndex = 8;
            this.lblSemestr.Text = "Семестр:";
            // 
            // cmbSemestrList
            // 
            this.cmbSemestrList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestrList.FormattingEnabled = true;
            this.cmbSemestrList.Location = new System.Drawing.Point(183, 65);
            this.cmbSemestrList.Name = "cmbSemestrList";
            this.cmbSemestrList.Size = new System.Drawing.Size(165, 21);
            this.cmbSemestrList.TabIndex = 9;
            // 
            // dtpClose
            // 
            this.dtpClose.Location = new System.Drawing.Point(183, 105);
            this.dtpClose.Name = "dtpClose";
            this.dtpClose.Size = new System.Drawing.Size(165, 20);
            this.dtpClose.TabIndex = 10;
            // 
            // dtpOpen
            // 
            this.dtpOpen.Location = new System.Drawing.Point(12, 105);
            this.dtpOpen.Name = "dtpOpen";
            this.dtpOpen.Size = new System.Drawing.Size(165, 20);
            this.dtpOpen.TabIndex = 11;
            // 
            // lblOpen
            // 
            this.lblOpen.AutoSize = true;
            this.lblOpen.Location = new System.Drawing.Point(12, 89);
            this.lblOpen.Name = "lblOpen";
            this.lblOpen.Size = new System.Drawing.Size(89, 13);
            this.lblOpen.TabIndex = 12;
            this.lblOpen.Text = "Открытие листа";
            // 
            // lblClose
            // 
            this.lblClose.AutoSize = true;
            this.lblClose.Location = new System.Drawing.Point(180, 89);
            this.lblClose.Name = "lblClose";
            this.lblClose.Size = new System.Drawing.Size(89, 13);
            this.lblClose.TabIndex = 13;
            this.lblClose.Text = "Закрытие листа";
            // 
            // lblDescript
            // 
            this.lblDescript.AutoSize = true;
            this.lblDescript.Location = new System.Drawing.Point(12, 128);
            this.lblDescript.Name = "lblDescript";
            this.lblDescript.Size = new System.Drawing.Size(70, 13);
            this.lblDescript.TabIndex = 14;
            this.lblDescript.Text = "Примечание";
            // 
            // txtDescript
            // 
            this.txtDescript.Location = new System.Drawing.Point(12, 144);
            this.txtDescript.Name = "txtDescript";
            this.txtDescript.Size = new System.Drawing.Size(336, 20);
            this.txtDescript.TabIndex = 15;
            // 
            // frmEditSickList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 262);
            this.Controls.Add(this.txtDescript);
            this.Controls.Add(this.lblDescript);
            this.Controls.Add(this.lblClose);
            this.Controls.Add(this.lblOpen);
            this.Controls.Add(this.dtpOpen);
            this.Controls.Add(this.dtpClose);
            this.Controls.Add(this.cmbSemestrList);
            this.Controls.Add(this.lblSemestr);
            this.Controls.Add(this.cmbLecturerList);
            this.Controls.Add(this.lblLecturer);
            this.Controls.Add(this.cmbLists);
            this.Controls.Add(this.lblLists);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditSickList";
            this.Text = "Редактирование больничных листов";
            this.Load += new System.EventHandler(this.frmEditSickList_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Label lblLists;
        private System.Windows.Forms.ComboBox cmbLists;
        private System.Windows.Forms.Label lblLecturer;
        private System.Windows.Forms.ComboBox cmbLecturerList;
        private System.Windows.Forms.Label lblSemestr;
        private System.Windows.Forms.ComboBox cmbSemestrList;
        private System.Windows.Forms.DateTimePicker dtpClose;
        private System.Windows.Forms.DateTimePicker dtpOpen;
        private System.Windows.Forms.Label lblOpen;
        private System.Windows.Forms.Label lblClose;
        private System.Windows.Forms.Label lblDescript;
        private System.Windows.Forms.TextBox txtDescript;
    }
}