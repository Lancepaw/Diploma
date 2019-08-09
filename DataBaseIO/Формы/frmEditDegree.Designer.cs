namespace DataBaseIO
{
    partial class frmEditDegree
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
            this.cmbDegreeList = new System.Windows.Forms.ComboBox();
            this.lblDegreeList = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDegree = new System.Windows.Forms.TextBox();
            this.lblShort = new System.Windows.Forms.Label();
            this.txtShort = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(168, 159);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(168, 130);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(87, 159);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 2;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(87, 130);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // cmbDegreeList
            // 
            this.cmbDegreeList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDegreeList.FormattingEnabled = true;
            this.cmbDegreeList.Location = new System.Drawing.Point(12, 25);
            this.cmbDegreeList.Name = "cmbDegreeList";
            this.cmbDegreeList.Size = new System.Drawing.Size(231, 21);
            this.cmbDegreeList.TabIndex = 4;
            this.cmbDegreeList.SelectedIndexChanged += new System.EventHandler(this.cmbDegreeList_SelectedIndexChanged);
            // 
            // lblDegreeList
            // 
            this.lblDegreeList.AutoSize = true;
            this.lblDegreeList.Location = new System.Drawing.Point(12, 9);
            this.lblDegreeList.Name = "lblDegreeList";
            this.lblDegreeList.Size = new System.Drawing.Size(49, 13);
            this.lblDegreeList.TabIndex = 5;
            this.lblDegreeList.Text = "Степени";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Степень";
            // 
            // txtDegree
            // 
            this.txtDegree.Location = new System.Drawing.Point(12, 65);
            this.txtDegree.Name = "txtDegree";
            this.txtDegree.Size = new System.Drawing.Size(231, 20);
            this.txtDegree.TabIndex = 7;
            // 
            // lblShort
            // 
            this.lblShort.AutoSize = true;
            this.lblShort.Location = new System.Drawing.Point(12, 88);
            this.lblShort.Name = "lblShort";
            this.lblShort.Size = new System.Drawing.Size(49, 13);
            this.lblShort.TabIndex = 8;
            this.lblShort.Text = "Коротко";
            // 
            // txtShort
            // 
            this.txtShort.Location = new System.Drawing.Point(12, 104);
            this.txtShort.Name = "txtShort";
            this.txtShort.Size = new System.Drawing.Size(231, 20);
            this.txtShort.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 164);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Всё готово";
            // 
            // frmEditDegree
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(255, 193);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtShort);
            this.Controls.Add(this.lblShort);
            this.Controls.Add(this.txtDegree);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblDegreeList);
            this.Controls.Add(this.cmbDegreeList);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditDegree";
            this.Text = "Редактор степеней";
            this.Load += new System.EventHandler(this.frmEditDegree_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.ComboBox cmbDegreeList;
        private System.Windows.Forms.Label lblDegreeList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDegree;
        private System.Windows.Forms.Label lblShort;
        private System.Windows.Forms.TextBox txtShort;
        private System.Windows.Forms.Label label2;
    }
}