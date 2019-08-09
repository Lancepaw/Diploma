namespace DataBaseIO
{
    partial class frmScheduleManagement
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
            this.dgScheduleManagement = new System.Windows.Forms.DataGridView();
            this.cmbSemestrList = new System.Windows.Forms.ComboBox();
            this.lblSemestrList = new System.Windows.Forms.Label();
            this.lblWorkYearList = new System.Windows.Forms.Label();
            this.cmbWorkYearList = new System.Windows.Forms.ComboBox();
            this.optMain = new System.Windows.Forms.RadioButton();
            this.optHoured = new System.Windows.Forms.RadioButton();
            this.optCombine = new System.Windows.Forms.RadioButton();
            this.optMainDop = new System.Windows.Forms.RadioButton();
            this.optCombineDop = new System.Windows.Forms.RadioButton();
            this.btnExcel = new System.Windows.Forms.Button();
            this.cmbForm = new System.Windows.Forms.ComboBox();
            this.chkPlan = new System.Windows.Forms.CheckBox();
            this.lblGrid = new System.Windows.Forms.Label();
            this.cmbGrid = new System.Windows.Forms.ComboBox();
            this.lblForm = new System.Windows.Forms.Label();
            this.btnForm = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgScheduleManagement)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(635, 446);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dgScheduleManagement
            // 
            this.dgScheduleManagement.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgScheduleManagement.Location = new System.Drawing.Point(12, 52);
            this.dgScheduleManagement.Name = "dgScheduleManagement";
            this.dgScheduleManagement.Size = new System.Drawing.Size(698, 345);
            this.dgScheduleManagement.TabIndex = 1;
            this.dgScheduleManagement.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgScheduleManagement_CellContentClick);
            // 
            // cmbSemestrList
            // 
            this.cmbSemestrList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSemestrList.FormattingEnabled = true;
            this.cmbSemestrList.Location = new System.Drawing.Point(12, 25);
            this.cmbSemestrList.Name = "cmbSemestrList";
            this.cmbSemestrList.Size = new System.Drawing.Size(121, 21);
            this.cmbSemestrList.TabIndex = 3;
            this.cmbSemestrList.SelectedIndexChanged += new System.EventHandler(this.cmbSemestrList_SelectedIndexChanged);
            // 
            // lblSemestrList
            // 
            this.lblSemestrList.AutoSize = true;
            this.lblSemestrList.Location = new System.Drawing.Point(12, 9);
            this.lblSemestrList.Name = "lblSemestrList";
            this.lblSemestrList.Size = new System.Drawing.Size(51, 13);
            this.lblSemestrList.TabIndex = 4;
            this.lblSemestrList.Text = "Семестр";
            // 
            // lblWorkYearList
            // 
            this.lblWorkYearList.AutoSize = true;
            this.lblWorkYearList.Location = new System.Drawing.Point(136, 9);
            this.lblWorkYearList.Name = "lblWorkYearList";
            this.lblWorkYearList.Size = new System.Drawing.Size(72, 13);
            this.lblWorkYearList.TabIndex = 5;
            this.lblWorkYearList.Text = "Учебный год";
            // 
            // cmbWorkYearList
            // 
            this.cmbWorkYearList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbWorkYearList.FormattingEnabled = true;
            this.cmbWorkYearList.Location = new System.Drawing.Point(139, 25);
            this.cmbWorkYearList.Name = "cmbWorkYearList";
            this.cmbWorkYearList.Size = new System.Drawing.Size(121, 21);
            this.cmbWorkYearList.TabIndex = 6;
            // 
            // optMain
            // 
            this.optMain.AutoSize = true;
            this.optMain.Location = new System.Drawing.Point(266, 7);
            this.optMain.Name = "optMain";
            this.optMain.Size = new System.Drawing.Size(67, 17);
            this.optMain.TabIndex = 9;
            this.optMain.TabStop = true;
            this.optMain.Text = "штатная";
            this.optMain.UseVisualStyleBackColor = true;
            this.optMain.CheckedChanged += new System.EventHandler(this.optMain_CheckedChanged);
            // 
            // optHoured
            // 
            this.optHoured.AutoSize = true;
            this.optHoured.Location = new System.Drawing.Point(490, 7);
            this.optHoured.Name = "optHoured";
            this.optHoured.Size = new System.Drawing.Size(78, 17);
            this.optHoured.TabIndex = 10;
            this.optHoured.TabStop = true;
            this.optHoured.Text = "почасовая";
            this.optHoured.UseVisualStyleBackColor = true;
            this.optHoured.CheckedChanged += new System.EventHandler(this.optHoured_CheckedChanged);
            // 
            // optCombine
            // 
            this.optCombine.AutoSize = true;
            this.optCombine.Location = new System.Drawing.Point(352, 7);
            this.optCombine.Name = "optCombine";
            this.optCombine.Size = new System.Drawing.Size(117, 17);
            this.optCombine.TabIndex = 11;
            this.optCombine.TabStop = true;
            this.optCombine.Text = "комбинированная";
            this.optCombine.UseVisualStyleBackColor = true;
            this.optCombine.CheckedChanged += new System.EventHandler(this.optCombine_CheckedChanged);
            // 
            // optMainDop
            // 
            this.optMainDop.AutoSize = true;
            this.optMainDop.Location = new System.Drawing.Point(266, 26);
            this.optMainDop.Name = "optMainDop";
            this.optMainDop.Size = new System.Drawing.Size(82, 17);
            this.optMainDop.TabIndex = 12;
            this.optMainDop.TabStop = true;
            this.optMainDop.Text = "штатная +0";
            this.optMainDop.UseVisualStyleBackColor = true;
            this.optMainDop.CheckedChanged += new System.EventHandler(this.optMainDop_CheckedChanged);
            // 
            // optCombineDop
            // 
            this.optCombineDop.AutoSize = true;
            this.optCombineDop.Location = new System.Drawing.Point(352, 26);
            this.optCombineDop.Name = "optCombineDop";
            this.optCombineDop.Size = new System.Drawing.Size(132, 17);
            this.optCombineDop.TabIndex = 13;
            this.optCombineDop.TabStop = true;
            this.optCombineDop.Text = "комбинированная +0";
            this.optCombineDop.UseVisualStyleBackColor = true;
            this.optCombineDop.CheckedChanged += new System.EventHandler(this.optCombineDop_CheckedChanged);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(541, 446);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(88, 23);
            this.btnExcel.TabIndex = 14;
            this.btnExcel.Text = "Сетку в Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // cmbForm
            // 
            this.cmbForm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbForm.FormattingEnabled = true;
            this.cmbForm.Location = new System.Drawing.Point(266, 419);
            this.cmbForm.Name = "cmbForm";
            this.cmbForm.Size = new System.Drawing.Size(218, 21);
            this.cmbForm.TabIndex = 17;
            this.cmbForm.SelectedIndexChanged += new System.EventHandler(this.cmbForm_SelectedIndexChanged);
            // 
            // chkPlan
            // 
            this.chkPlan.AutoSize = true;
            this.chkPlan.Location = new System.Drawing.Point(490, 27);
            this.chkPlan.Name = "chkPlan";
            this.chkPlan.Size = new System.Drawing.Size(76, 17);
            this.chkPlan.TabIndex = 19;
            this.chkPlan.Text = "Плановая";
            this.chkPlan.UseVisualStyleBackColor = true;
            // 
            // lblGrid
            // 
            this.lblGrid.AutoSize = true;
            this.lblGrid.Location = new System.Drawing.Point(12, 403);
            this.lblGrid.Name = "lblGrid";
            this.lblGrid.Size = new System.Drawing.Size(142, 13);
            this.lblGrid.TabIndex = 21;
            this.lblGrid.Text = "В сетку (экранная форма):";
            // 
            // cmbGrid
            // 
            this.cmbGrid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbGrid.FormattingEnabled = true;
            this.cmbGrid.Location = new System.Drawing.Point(12, 419);
            this.cmbGrid.Name = "cmbGrid";
            this.cmbGrid.Size = new System.Drawing.Size(248, 21);
            this.cmbGrid.TabIndex = 22;
            this.cmbGrid.SelectedIndexChanged += new System.EventHandler(this.cmbGrid_SelectedIndexChanged);
            // 
            // lblForm
            // 
            this.lblForm.AutoSize = true;
            this.lblForm.Location = new System.Drawing.Point(263, 400);
            this.lblForm.Name = "lblForm";
            this.lblForm.Size = new System.Drawing.Size(149, 13);
            this.lblForm.TabIndex = 23;
            this.lblForm.Text = "Документ (отчётная форма)";
            // 
            // btnForm
            // 
            this.btnForm.Location = new System.Drawing.Point(266, 446);
            this.btnForm.Name = "btnForm";
            this.btnForm.Size = new System.Drawing.Size(95, 23);
            this.btnForm.TabIndex = 24;
            this.btnForm.Text = "Сформировать";
            this.btnForm.UseVisualStyleBackColor = true;
            this.btnForm.Click += new System.EventHandler(this.btnForm_Click);
            // 
            // frmScheduleManagement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(722, 475);
            this.Controls.Add(this.btnForm);
            this.Controls.Add(this.lblForm);
            this.Controls.Add(this.cmbGrid);
            this.Controls.Add(this.lblGrid);
            this.Controls.Add(this.chkPlan);
            this.Controls.Add(this.cmbForm);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.optCombineDop);
            this.Controls.Add(this.optMainDop);
            this.Controls.Add(this.optCombine);
            this.Controls.Add(this.optHoured);
            this.Controls.Add(this.optMain);
            this.Controls.Add(this.cmbWorkYearList);
            this.Controls.Add(this.lblWorkYearList);
            this.Controls.Add(this.lblSemestrList);
            this.Controls.Add(this.cmbSemestrList);
            this.Controls.Add(this.dgScheduleManagement);
            this.Controls.Add(this.btnClose);
            this.Name = "frmScheduleManagement";
            this.Text = "Формирование заявки в диспетчерскую";
            this.Load += new System.EventHandler(this.frmScheduleManagement_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgScheduleManagement)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.DataGridView dgScheduleManagement;
        private System.Windows.Forms.ComboBox cmbSemestrList;
        private System.Windows.Forms.Label lblSemestrList;
        private System.Windows.Forms.Label lblWorkYearList;
        private System.Windows.Forms.ComboBox cmbWorkYearList;
        private System.Windows.Forms.RadioButton optMain;
        private System.Windows.Forms.RadioButton optHoured;
        private System.Windows.Forms.RadioButton optCombine;
        private System.Windows.Forms.RadioButton optMainDop;
        private System.Windows.Forms.RadioButton optCombineDop;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.ComboBox cmbForm;
        private System.Windows.Forms.CheckBox chkPlan;
        private System.Windows.Forms.Label lblGrid;
        private System.Windows.Forms.ComboBox cmbGrid;
        private System.Windows.Forms.Label lblForm;
        private System.Windows.Forms.Button btnForm;
    }
}