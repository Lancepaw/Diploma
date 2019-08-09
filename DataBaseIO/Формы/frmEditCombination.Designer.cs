namespace DataBaseIO
{
    partial class frmEditCombination
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
            this.cmbCombinationList = new System.Windows.Forms.ComboBox();
            this.lblCombinationList = new System.Windows.Forms.Label();
            this.lblCombination = new System.Windows.Forms.Label();
            this.txtCombination = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(202, 120);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(202, 91);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(121, 91);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(121, 120);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // cmbCombinationList
            // 
            this.cmbCombinationList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCombinationList.FormattingEnabled = true;
            this.cmbCombinationList.Location = new System.Drawing.Point(12, 25);
            this.cmbCombinationList.Name = "cmbCombinationList";
            this.cmbCombinationList.Size = new System.Drawing.Size(265, 21);
            this.cmbCombinationList.TabIndex = 4;
            this.cmbCombinationList.SelectedIndexChanged += new System.EventHandler(this.cmbCombinationList_SelectedIndexChanged);
            // 
            // lblCombinationList
            // 
            this.lblCombinationList.AutoSize = true;
            this.lblCombinationList.Location = new System.Drawing.Point(12, 9);
            this.lblCombinationList.Name = "lblCombinationList";
            this.lblCombinationList.Size = new System.Drawing.Size(106, 13);
            this.lblCombinationList.TabIndex = 5;
            this.lblCombinationList.Text = "Совместительство:";
            // 
            // lblCombination
            // 
            this.lblCombination.AutoSize = true;
            this.lblCombination.Location = new System.Drawing.Point(12, 49);
            this.lblCombination.Name = "lblCombination";
            this.lblCombination.Size = new System.Drawing.Size(127, 13);
            this.lblCombination.TabIndex = 6;
            this.lblCombination.Text = "Тип совместительства:";
            // 
            // txtCombination
            // 
            this.txtCombination.Location = new System.Drawing.Point(12, 65);
            this.txtCombination.Name = "txtCombination";
            this.txtCombination.Size = new System.Drawing.Size(265, 20);
            this.txtCombination.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 125);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Всё готово";
            // 
            // frmEditCombination
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(289, 154);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCombination);
            this.Controls.Add(this.lblCombination);
            this.Controls.Add(this.lblCombinationList);
            this.Controls.Add(this.cmbCombinationList);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Name = "frmEditCombination";
            this.Text = "Редактор совместительства";
            this.Load += new System.EventHandler(this.frmEditCombination_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.ComboBox cmbCombinationList;
        private System.Windows.Forms.Label lblCombinationList;
        private System.Windows.Forms.Label lblCombination;
        private System.Windows.Forms.TextBox txtCombination;
        private System.Windows.Forms.Label label1;
    }
}