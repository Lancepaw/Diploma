namespace DataBaseIO
{
    partial class frmSelectLect
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
            this.lblChoose = new System.Windows.Forms.Label();
            this.cmbChoose = new System.Windows.Forms.ComboBox();
            this.btnHor = new System.Windows.Forms.Button();
            this.btnVert = new System.Windows.Forms.Button();
            this.btnAll = new System.Windows.Forms.Button();
            this.lblInOne = new System.Windows.Forms.Label();
            this.lblindOneVert = new System.Windows.Forms.Label();
            this.lblindOneHor = new System.Windows.Forms.Label();
            this.lblforAll = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblChoose
            // 
            this.lblChoose.AutoSize = true;
            this.lblChoose.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblChoose.Location = new System.Drawing.Point(56, 15);
            this.lblChoose.Name = "lblChoose";
            this.lblChoose.Size = new System.Drawing.Size(275, 25);
            this.lblChoose.TabIndex = 0;
            this.lblChoose.Text = "Выберите преподавателя:";
            // 
            // cmbChoose
            // 
            this.cmbChoose.FormattingEnabled = true;
            this.cmbChoose.Location = new System.Drawing.Point(12, 57);
            this.cmbChoose.Name = "cmbChoose";
            this.cmbChoose.Size = new System.Drawing.Size(360, 21);
            this.cmbChoose.TabIndex = 1;
            // 
            // btnHor
            // 
            this.btnHor.Location = new System.Drawing.Point(257, 158);
            this.btnHor.Name = "btnHor";
            this.btnHor.Size = new System.Drawing.Size(115, 23);
            this.btnHor.TabIndex = 2;
            this.btnHor.Text = "Формировать";
            this.btnHor.UseVisualStyleBackColor = true;
            this.btnHor.Click += new System.EventHandler(this.btnChoose_Click);
            // 
            // btnVert
            // 
            this.btnVert.Location = new System.Drawing.Point(257, 129);
            this.btnVert.Name = "btnVert";
            this.btnVert.Size = new System.Drawing.Size(115, 23);
            this.btnVert.TabIndex = 3;
            this.btnVert.Text = "Формировать";
            this.btnVert.UseVisualStyleBackColor = true;
            this.btnVert.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnAll
            // 
            this.btnAll.Location = new System.Drawing.Point(118, 226);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(145, 23);
            this.btnAll.TabIndex = 4;
            this.btnAll.Text = "Создать и сохранить";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // lblInOne
            // 
            this.lblInOne.AutoSize = true;
            this.lblInOne.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblInOne.Location = new System.Drawing.Point(12, 101);
            this.lblInOne.Name = "lblInOne";
            this.lblInOne.Size = new System.Drawing.Size(366, 15);
            this.lblInOne.TabIndex = 5;
            this.lblInOne.Text = "Формирование индивидуального расписания преподавателя";
            // 
            // lblindOneVert
            // 
            this.lblindOneVert.AutoSize = true;
            this.lblindOneVert.Location = new System.Drawing.Point(12, 134);
            this.lblindOneVert.Name = "lblindOneVert";
            this.lblindOneVert.Size = new System.Drawing.Size(201, 13);
            this.lblindOneVert.TabIndex = 6;
            this.lblindOneVert.Text = "Расписание в вертикальном формате";
            // 
            // lblindOneHor
            // 
            this.lblindOneHor.AutoSize = true;
            this.lblindOneHor.Location = new System.Drawing.Point(12, 158);
            this.lblindOneHor.Name = "lblindOneHor";
            this.lblindOneHor.Size = new System.Drawing.Size(212, 13);
            this.lblindOneHor.TabIndex = 7;
            this.lblindOneHor.Text = "Расписание в горизонтальном формате";
            // 
            // lblforAll
            // 
            this.lblforAll.AutoSize = true;
            this.lblforAll.Location = new System.Drawing.Point(18, 197);
            this.lblforAll.Name = "lblforAll";
            this.lblforAll.Size = new System.Drawing.Size(354, 13);
            this.lblforAll.TabIndex = 8;
            this.lblforAll.Text = "Сохранение карточек расписаний каждого преподавателя кафедры";
            // 
            // frmSelectLect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 261);
            this.Controls.Add(this.lblforAll);
            this.Controls.Add(this.lblindOneHor);
            this.Controls.Add(this.lblindOneVert);
            this.Controls.Add(this.lblInOne);
            this.Controls.Add(this.btnAll);
            this.Controls.Add(this.btnVert);
            this.Controls.Add(this.btnHor);
            this.Controls.Add(this.cmbChoose);
            this.Controls.Add(this.lblChoose);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(400, 300);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 175);
            this.Name = "frmSelectLect";
            this.Text = "Выберите преподавателя";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblChoose;
        private System.Windows.Forms.ComboBox cmbChoose;
        private System.Windows.Forms.Button btnHor;
        private System.Windows.Forms.Button btnVert;
        private System.Windows.Forms.Button btnAll;
        private System.Windows.Forms.Label lblInOne;
        private System.Windows.Forms.Label lblindOneVert;
        private System.Windows.Forms.Label lblindOneHor;
        private System.Windows.Forms.Label lblforAll;
    }
}