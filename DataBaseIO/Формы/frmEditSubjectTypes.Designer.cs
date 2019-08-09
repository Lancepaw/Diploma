namespace DataBaseIO
{
    partial class frmEditSubjectTypes
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
            this.lblSubjectTypesList = new System.Windows.Forms.Label();
            this.cmbSubjectTypesList = new System.Windows.Forms.ComboBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lblType = new System.Windows.Forms.Label();
            this.txtType = new System.Windows.Forms.TextBox();
            this.lblShort = new System.Windows.Forms.Label();
            this.txtShort = new System.Windows.Forms.TextBox();
            this.lblPlan = new System.Windows.Forms.Label();
            this.txtPlan = new System.Windows.Forms.TextBox();
            this.lblDistrib = new System.Windows.Forms.Label();
            this.txtDistrib = new System.Windows.Forms.TextBox();
            this.lblForms = new System.Windows.Forms.Label();
            this.txtForms = new System.Windows.Forms.TextBox();
            this.btnMoveUp = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblSubjectTypesList
            // 
            this.lblSubjectTypesList.AutoSize = true;
            this.lblSubjectTypesList.Location = new System.Drawing.Point(12, 9);
            this.lblSubjectTypesList.Name = "lblSubjectTypesList";
            this.lblSubjectTypesList.Size = new System.Drawing.Size(181, 13);
            this.lblSubjectTypesList.TabIndex = 0;
            this.lblSubjectTypesList.Text = "Перечень видов учебной нагрузки";
            // 
            // cmbSubjectTypesList
            // 
            this.cmbSubjectTypesList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSubjectTypesList.FormattingEnabled = true;
            this.cmbSubjectTypesList.Location = new System.Drawing.Point(12, 25);
            this.cmbSubjectTypesList.Name = "cmbSubjectTypesList";
            this.cmbSubjectTypesList.Size = new System.Drawing.Size(405, 21);
            this.cmbSubjectTypesList.TabIndex = 1;
            this.cmbSubjectTypesList.SelectedIndexChanged += new System.EventHandler(this.cmbSubjectTypesList_SelectedIndexChanged);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(342, 238);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(261, 238);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(342, 209);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(261, 209);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lblType
            // 
            this.lblType.AutoSize = true;
            this.lblType.Location = new System.Drawing.Point(12, 49);
            this.lblType.Name = "lblType";
            this.lblType.Size = new System.Drawing.Size(118, 13);
            this.lblType.TabIndex = 6;
            this.lblType.Text = "Вид учебной нагрузки";
            // 
            // txtType
            // 
            this.txtType.Location = new System.Drawing.Point(12, 65);
            this.txtType.Name = "txtType";
            this.txtType.Size = new System.Drawing.Size(118, 20);
            this.txtType.TabIndex = 7;
            // 
            // lblShort
            // 
            this.lblShort.AutoSize = true;
            this.lblShort.Location = new System.Drawing.Point(136, 49);
            this.lblShort.Name = "lblShort";
            this.lblShort.Size = new System.Drawing.Size(132, 13);
            this.lblShort.TabIndex = 8;
            this.lblShort.Text = "Короткое наименование";
            // 
            // txtShort
            // 
            this.txtShort.Location = new System.Drawing.Point(139, 65);
            this.txtShort.Name = "txtShort";
            this.txtShort.Size = new System.Drawing.Size(129, 20);
            this.txtShort.TabIndex = 9;
            // 
            // lblPlan
            // 
            this.lblPlan.AutoSize = true;
            this.lblPlan.Location = new System.Drawing.Point(274, 49);
            this.lblPlan.Name = "lblPlan";
            this.lblPlan.Size = new System.Drawing.Size(129, 13);
            this.lblPlan.TabIndex = 10;
            this.lblPlan.Text = "В индивидуальный план";
            // 
            // txtPlan
            // 
            this.txtPlan.Location = new System.Drawing.Point(277, 65);
            this.txtPlan.Name = "txtPlan";
            this.txtPlan.Size = new System.Drawing.Size(140, 20);
            this.txtPlan.TabIndex = 11;
            // 
            // lblDistrib
            // 
            this.lblDistrib.AutoSize = true;
            this.lblDistrib.Location = new System.Drawing.Point(12, 88);
            this.lblDistrib.Name = "lblDistrib";
            this.lblDistrib.Size = new System.Drawing.Size(160, 13);
            this.lblDistrib.TabIndex = 12;
            this.lblDistrib.Text = "Как в таблице распределения";
            // 
            // txtDistrib
            // 
            this.txtDistrib.Location = new System.Drawing.Point(12, 104);
            this.txtDistrib.Name = "txtDistrib";
            this.txtDistrib.Size = new System.Drawing.Size(160, 20);
            this.txtDistrib.TabIndex = 13;
            // 
            // lblForms
            // 
            this.lblForms.AutoSize = true;
            this.lblForms.Location = new System.Drawing.Point(178, 88);
            this.lblForms.Name = "lblForms";
            this.lblForms.Size = new System.Drawing.Size(109, 13);
            this.lblForms.TabIndex = 14;
            this.lblForms.Text = "Для печатных форм";
            // 
            // txtForms
            // 
            this.txtForms.Location = new System.Drawing.Point(181, 104);
            this.txtForms.Name = "txtForms";
            this.txtForms.Size = new System.Drawing.Size(106, 20);
            this.txtForms.TabIndex = 15;
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Location = new System.Drawing.Point(12, 130);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(49, 23);
            this.btnMoveUp.TabIndex = 16;
            this.btnMoveUp.Text = "Вверх";
            this.btnMoveUp.UseVisualStyleBackColor = true;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(12, 159);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(49, 23);
            this.btnDown.TabIndex = 17;
            this.btnDown.Text = "Вниз";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // frmEditSubjectTypes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(429, 273);
            this.Controls.Add(this.btnDown);
            this.Controls.Add(this.btnMoveUp);
            this.Controls.Add(this.txtForms);
            this.Controls.Add(this.lblForms);
            this.Controls.Add(this.txtDistrib);
            this.Controls.Add(this.lblDistrib);
            this.Controls.Add(this.txtPlan);
            this.Controls.Add(this.lblPlan);
            this.Controls.Add(this.txtShort);
            this.Controls.Add(this.lblShort);
            this.Controls.Add(this.txtType);
            this.Controls.Add(this.lblType);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.cmbSubjectTypesList);
            this.Controls.Add(this.lblSubjectTypesList);
            this.Name = "frmEditSubjectTypes";
            this.Text = "Редактирование видов учебной нагрузки";
            this.Load += new System.EventHandler(this.frmEditSubjectTypes_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSubjectTypesList;
        private System.Windows.Forms.ComboBox cmbSubjectTypesList;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label lblType;
        private System.Windows.Forms.TextBox txtType;
        private System.Windows.Forms.Label lblShort;
        private System.Windows.Forms.TextBox txtShort;
        private System.Windows.Forms.Label lblPlan;
        private System.Windows.Forms.TextBox txtPlan;
        private System.Windows.Forms.Label lblDistrib;
        private System.Windows.Forms.TextBox txtDistrib;
        private System.Windows.Forms.Label lblForms;
        private System.Windows.Forms.TextBox txtForms;
        private System.Windows.Forms.Button btnMoveUp;
        private System.Windows.Forms.Button btnDown;
    }
}