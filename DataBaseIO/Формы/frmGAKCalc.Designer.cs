namespace DataBaseIO
{
    partial class frmGAKCalc
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
            this.lblGeneral = new System.Windows.Forms.Label();
            this.txtGeneral = new System.Windows.Forms.TextBox();
            this.lblStudNum = new System.Windows.Forms.Label();
            this.txtStudNum = new System.Windows.Forms.TextBox();
            this.lblStudentLoad = new System.Windows.Forms.Label();
            this.txtStudentLoad = new System.Windows.Forms.TextBox();
            this.btnCount = new System.Windows.Forms.Button();
            this.lblDiplomaLoad = new System.Windows.Forms.Label();
            this.lblGAK = new System.Windows.Forms.Label();
            this.txtDiploma = new System.Windows.Forms.TextBox();
            this.txtGAK = new System.Windows.Forms.TextBox();
            this.lblPeople = new System.Windows.Forms.Label();
            this.txtPeople = new System.Windows.Forms.TextBox();
            this.lblMainInGAK = new System.Windows.Forms.Label();
            this.txtToMain = new System.Windows.Forms.TextBox();
            this.lblLect = new System.Windows.Forms.Label();
            this.txtLect = new System.Windows.Forms.TextBox();
            this.lblOuter = new System.Windows.Forms.Label();
            this.txtOuter = new System.Windows.Forms.TextBox();
            this.lblGeneralScheme = new System.Windows.Forms.Label();
            this.lblGen = new System.Windows.Forms.Label();
            this.lblInner = new System.Windows.Forms.Label();
            this.txtInner = new System.Windows.Forms.TextBox();
            this.lblList = new System.Windows.Forms.Label();
            this.cmbDiplomaList = new System.Windows.Forms.ComboBox();
            this.lblSubscribe = new System.Windows.Forms.Label();
            this.txtSubscribe = new System.Windows.Forms.TextBox();
            this.txtSecr = new System.Windows.Forms.TextBox();
            this.lblSecr = new System.Windows.Forms.Label();
            this.lblToOuter = new System.Windows.Forms.Label();
            this.txtToOuter = new System.Windows.Forms.TextBox();
            this.txtToInner = new System.Windows.Forms.TextBox();
            this.lblToInner = new System.Windows.Forms.Label();
            this.txtMain = new System.Windows.Forms.TextBox();
            this.lblMain = new System.Windows.Forms.Label();
            this.txtToSecretary = new System.Windows.Forms.TextBox();
            this.lblToSecretary = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(335, 509);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblGeneral
            // 
            this.lblGeneral.AutoSize = true;
            this.lblGeneral.Location = new System.Drawing.Point(13, 68);
            this.lblGeneral.Name = "lblGeneral";
            this.lblGeneral.Size = new System.Drawing.Size(135, 13);
            this.lblGeneral.TabIndex = 1;
            this.lblGeneral.Text = "Общее количество часов";
            // 
            // txtGeneral
            // 
            this.txtGeneral.Location = new System.Drawing.Point(192, 65);
            this.txtGeneral.Name = "txtGeneral";
            this.txtGeneral.Size = new System.Drawing.Size(62, 20);
            this.txtGeneral.TabIndex = 2;
            // 
            // lblStudNum
            // 
            this.lblStudNum.AutoSize = true;
            this.lblStudNum.Location = new System.Drawing.Point(13, 94);
            this.lblStudNum.Name = "lblStudNum";
            this.lblStudNum.Size = new System.Drawing.Size(120, 13);
            this.lblStudNum.TabIndex = 3;
            this.lblStudNum.Text = "Количество студентов";
            // 
            // txtStudNum
            // 
            this.txtStudNum.Location = new System.Drawing.Point(192, 91);
            this.txtStudNum.Name = "txtStudNum";
            this.txtStudNum.Size = new System.Drawing.Size(62, 20);
            this.txtStudNum.TabIndex = 4;
            // 
            // lblStudentLoad
            // 
            this.lblStudentLoad.AutoSize = true;
            this.lblStudentLoad.Location = new System.Drawing.Point(13, 270);
            this.lblStudentLoad.Name = "lblStudentLoad";
            this.lblStudentLoad.Size = new System.Drawing.Size(173, 13);
            this.lblStudentLoad.TabIndex = 5;
            this.lblStudentLoad.Text = "Нагрузка на одного дипломника";
            // 
            // txtStudentLoad
            // 
            this.txtStudentLoad.Location = new System.Drawing.Point(192, 267);
            this.txtStudentLoad.Name = "txtStudentLoad";
            this.txtStudentLoad.Size = new System.Drawing.Size(62, 20);
            this.txtStudentLoad.TabIndex = 6;
            // 
            // btnCount
            // 
            this.btnCount.Location = new System.Drawing.Point(12, 509);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 23);
            this.btnCount.TabIndex = 7;
            this.btnCount.Text = "Расчитать";
            this.btnCount.UseVisualStyleBackColor = true;
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // lblDiplomaLoad
            // 
            this.lblDiplomaLoad.AutoSize = true;
            this.lblDiplomaLoad.Location = new System.Drawing.Point(13, 296);
            this.lblDiplomaLoad.Name = "lblDiplomaLoad";
            this.lblDiplomaLoad.Size = new System.Drawing.Size(166, 13);
            this.lblDiplomaLoad.TabIndex = 8;
            this.lblDiplomaLoad.Text = "На дипломное проектирование";
            // 
            // lblGAK
            // 
            this.lblGAK.AutoSize = true;
            this.lblGAK.Location = new System.Drawing.Point(13, 322);
            this.lblGAK.Name = "lblGAK";
            this.lblGAK.Size = new System.Drawing.Size(44, 13);
            this.lblGAK.TabIndex = 9;
            this.lblGAK.Text = "На ГАК";
            // 
            // txtDiploma
            // 
            this.txtDiploma.Location = new System.Drawing.Point(192, 293);
            this.txtDiploma.Name = "txtDiploma";
            this.txtDiploma.Size = new System.Drawing.Size(62, 20);
            this.txtDiploma.TabIndex = 10;
            // 
            // txtGAK
            // 
            this.txtGAK.Location = new System.Drawing.Point(192, 319);
            this.txtGAK.Name = "txtGAK";
            this.txtGAK.Size = new System.Drawing.Size(62, 20);
            this.txtGAK.TabIndex = 11;
            // 
            // lblPeople
            // 
            this.lblPeople.AutoSize = true;
            this.lblPeople.Location = new System.Drawing.Point(13, 348);
            this.lblPeople.Name = "lblPeople";
            this.lblPeople.Size = new System.Drawing.Size(97, 13);
            this.lblPeople.TabIndex = 12;
            this.lblPeople.Text = "Численность ГАК";
            // 
            // txtPeople
            // 
            this.txtPeople.Location = new System.Drawing.Point(192, 345);
            this.txtPeople.Name = "txtPeople";
            this.txtPeople.Size = new System.Drawing.Size(62, 20);
            this.txtPeople.TabIndex = 13;
            // 
            // lblMainInGAK
            // 
            this.lblMainInGAK.AutoSize = true;
            this.lblMainInGAK.Location = new System.Drawing.Point(13, 146);
            this.lblMainInGAK.Name = "lblMainInGAK";
            this.lblMainInGAK.Size = new System.Drawing.Size(82, 13);
            this.lblMainInGAK.TabIndex = 14;
            this.lblMainInGAK.Text = "Председателю";
            // 
            // txtToMain
            // 
            this.txtToMain.Location = new System.Drawing.Point(192, 143);
            this.txtToMain.Name = "txtToMain";
            this.txtToMain.Size = new System.Drawing.Size(62, 20);
            this.txtToMain.TabIndex = 15;
            // 
            // lblLect
            // 
            this.lblLect.AutoSize = true;
            this.lblLect.Location = new System.Drawing.Point(13, 478);
            this.lblLect.Name = "lblLect";
            this.lblLect.Size = new System.Drawing.Size(94, 13);
            this.lblLect.TabIndex = 16;
            this.lblLect.Text = "Преподавателям";
            // 
            // txtLect
            // 
            this.txtLect.Location = new System.Drawing.Point(192, 475);
            this.txtLect.Name = "txtLect";
            this.txtLect.Size = new System.Drawing.Size(62, 20);
            this.txtLect.TabIndex = 17;
            // 
            // lblOuter
            // 
            this.lblOuter.AutoSize = true;
            this.lblOuter.Location = new System.Drawing.Point(135, 400);
            this.lblOuter.Name = "lblOuter";
            this.lblOuter.Size = new System.Drawing.Size(51, 13);
            this.lblOuter.TabIndex = 18;
            this.lblOuter.Text = "внешние";
            // 
            // txtOuter
            // 
            this.txtOuter.Location = new System.Drawing.Point(192, 397);
            this.txtOuter.Name = "txtOuter";
            this.txtOuter.Size = new System.Drawing.Size(62, 20);
            this.txtOuter.TabIndex = 19;
            // 
            // lblGeneralScheme
            // 
            this.lblGeneralScheme.AutoSize = true;
            this.lblGeneralScheme.Location = new System.Drawing.Point(13, 46);
            this.lblGeneralScheme.Name = "lblGeneralScheme";
            this.lblGeneralScheme.Size = new System.Drawing.Size(79, 13);
            this.lblGeneralScheme.TabIndex = 20;
            this.lblGeneralScheme.Text = "Общая схема:";
            // 
            // lblGen
            // 
            this.lblGen.AutoSize = true;
            this.lblGen.Location = new System.Drawing.Point(146, 348);
            this.lblGen.Name = "lblGen";
            this.lblGen.Size = new System.Drawing.Size(40, 13);
            this.lblGen.TabIndex = 24;
            this.lblGen.Text = "общее";
            // 
            // lblInner
            // 
            this.lblInner.AutoSize = true;
            this.lblInner.Location = new System.Drawing.Point(121, 426);
            this.lblInner.Name = "lblInner";
            this.lblInner.Size = new System.Drawing.Size(65, 13);
            this.lblInner.TabIndex = 25;
            this.lblInner.Text = "внутренние";
            // 
            // txtInner
            // 
            this.txtInner.Location = new System.Drawing.Point(192, 423);
            this.txtInner.Name = "txtInner";
            this.txtInner.Size = new System.Drawing.Size(62, 20);
            this.txtInner.TabIndex = 26;
            // 
            // lblList
            // 
            this.lblList.AutoSize = true;
            this.lblList.Location = new System.Drawing.Point(13, 18);
            this.lblList.Name = "lblList";
            this.lblList.Size = new System.Drawing.Size(197, 13);
            this.lblList.TabIndex = 27;
            this.lblList.Text = "Список дипломного проектирования:";
            // 
            // cmbDiplomaList
            // 
            this.cmbDiplomaList.FormattingEnabled = true;
            this.cmbDiplomaList.Location = new System.Drawing.Point(216, 15);
            this.cmbDiplomaList.Name = "cmbDiplomaList";
            this.cmbDiplomaList.Size = new System.Drawing.Size(200, 21);
            this.cmbDiplomaList.TabIndex = 28;
            this.cmbDiplomaList.SelectedIndexChanged += new System.EventHandler(this.cmbDiplomaList_SelectedIndexChanged);
            // 
            // lblSubscribe
            // 
            this.lblSubscribe.AutoSize = true;
            this.lblSubscribe.Location = new System.Drawing.Point(13, 120);
            this.lblSubscribe.Name = "lblSubscribe";
            this.lblSubscribe.Size = new System.Drawing.Size(161, 13);
            this.lblSubscribe.TabIndex = 29;
            this.lblSubscribe.Text = "Подпись дипломных проектов";
            // 
            // txtSubscribe
            // 
            this.txtSubscribe.Location = new System.Drawing.Point(192, 117);
            this.txtSubscribe.Name = "txtSubscribe";
            this.txtSubscribe.Size = new System.Drawing.Size(62, 20);
            this.txtSubscribe.TabIndex = 30;
            // 
            // txtSecr
            // 
            this.txtSecr.Location = new System.Drawing.Point(192, 449);
            this.txtSecr.Name = "txtSecr";
            this.txtSecr.Size = new System.Drawing.Size(62, 20);
            this.txtSecr.TabIndex = 31;
            // 
            // lblSecr
            // 
            this.lblSecr.AutoSize = true;
            this.lblSecr.Location = new System.Drawing.Point(126, 452);
            this.lblSecr.Name = "lblSecr";
            this.lblSecr.Size = new System.Drawing.Size(60, 13);
            this.lblSecr.TabIndex = 32;
            this.lblSecr.Text = "секретарь";
            // 
            // lblToOuter
            // 
            this.lblToOuter.AutoSize = true;
            this.lblToOuter.Location = new System.Drawing.Point(13, 172);
            this.lblToOuter.Name = "lblToOuter";
            this.lblToOuter.Size = new System.Drawing.Size(54, 13);
            this.lblToOuter.TabIndex = 33;
            this.lblToOuter.Text = "Внешним";
            // 
            // txtToOuter
            // 
            this.txtToOuter.Location = new System.Drawing.Point(192, 169);
            this.txtToOuter.Name = "txtToOuter";
            this.txtToOuter.Size = new System.Drawing.Size(62, 20);
            this.txtToOuter.TabIndex = 34;
            // 
            // txtToInner
            // 
            this.txtToInner.Location = new System.Drawing.Point(192, 195);
            this.txtToInner.Name = "txtToInner";
            this.txtToInner.Size = new System.Drawing.Size(62, 20);
            this.txtToInner.TabIndex = 35;
            // 
            // lblToInner
            // 
            this.lblToInner.AutoSize = true;
            this.lblToInner.Location = new System.Drawing.Point(13, 198);
            this.lblToInner.Name = "lblToInner";
            this.lblToInner.Size = new System.Drawing.Size(68, 13);
            this.lblToInner.TabIndex = 36;
            this.lblToInner.Text = "Внутренним";
            // 
            // txtMain
            // 
            this.txtMain.Location = new System.Drawing.Point(192, 371);
            this.txtMain.Name = "txtMain";
            this.txtMain.Size = new System.Drawing.Size(62, 20);
            this.txtMain.TabIndex = 37;
            // 
            // lblMain
            // 
            this.lblMain.AutoSize = true;
            this.lblMain.Location = new System.Drawing.Point(108, 374);
            this.lblMain.Name = "lblMain";
            this.lblMain.Size = new System.Drawing.Size(78, 13);
            this.lblMain.TabIndex = 38;
            this.lblMain.Text = "председатель";
            // 
            // txtToSecretary
            // 
            this.txtToSecretary.Location = new System.Drawing.Point(192, 221);
            this.txtToSecretary.Name = "txtToSecretary";
            this.txtToSecretary.Size = new System.Drawing.Size(62, 20);
            this.txtToSecretary.TabIndex = 39;
            // 
            // lblToSecretary
            // 
            this.lblToSecretary.AutoSize = true;
            this.lblToSecretary.Location = new System.Drawing.Point(13, 224);
            this.lblToSecretary.Name = "lblToSecretary";
            this.lblToSecretary.Size = new System.Drawing.Size(63, 13);
            this.lblToSecretary.TabIndex = 40;
            this.lblToSecretary.Text = "Секретарю";
            // 
            // frmGAKCalc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(422, 544);
            this.Controls.Add(this.lblToSecretary);
            this.Controls.Add(this.txtToSecretary);
            this.Controls.Add(this.lblMain);
            this.Controls.Add(this.txtMain);
            this.Controls.Add(this.lblToInner);
            this.Controls.Add(this.txtToInner);
            this.Controls.Add(this.txtToOuter);
            this.Controls.Add(this.lblToOuter);
            this.Controls.Add(this.lblSecr);
            this.Controls.Add(this.txtSecr);
            this.Controls.Add(this.txtSubscribe);
            this.Controls.Add(this.lblSubscribe);
            this.Controls.Add(this.cmbDiplomaList);
            this.Controls.Add(this.lblList);
            this.Controls.Add(this.txtInner);
            this.Controls.Add(this.lblInner);
            this.Controls.Add(this.lblGen);
            this.Controls.Add(this.lblGeneralScheme);
            this.Controls.Add(this.txtOuter);
            this.Controls.Add(this.lblOuter);
            this.Controls.Add(this.txtLect);
            this.Controls.Add(this.lblLect);
            this.Controls.Add(this.txtToMain);
            this.Controls.Add(this.lblMainInGAK);
            this.Controls.Add(this.txtPeople);
            this.Controls.Add(this.lblPeople);
            this.Controls.Add(this.txtGAK);
            this.Controls.Add(this.txtDiploma);
            this.Controls.Add(this.lblGAK);
            this.Controls.Add(this.lblDiplomaLoad);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.txtStudentLoad);
            this.Controls.Add(this.lblStudentLoad);
            this.Controls.Add(this.txtStudNum);
            this.Controls.Add(this.lblStudNum);
            this.Controls.Add(this.txtGeneral);
            this.Controls.Add(this.lblGeneral);
            this.Controls.Add(this.btnClose);
            this.Name = "frmGAKCalc";
            this.Text = "Калькулятор экзаменационной комиссии";
            this.Load += new System.EventHandler(this.frmGAKCalc_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblGeneral;
        private System.Windows.Forms.TextBox txtGeneral;
        private System.Windows.Forms.Label lblStudNum;
        private System.Windows.Forms.TextBox txtStudNum;
        private System.Windows.Forms.Label lblStudentLoad;
        private System.Windows.Forms.TextBox txtStudentLoad;
        private System.Windows.Forms.Button btnCount;
        private System.Windows.Forms.Label lblDiplomaLoad;
        private System.Windows.Forms.Label lblGAK;
        private System.Windows.Forms.TextBox txtDiploma;
        private System.Windows.Forms.TextBox txtGAK;
        private System.Windows.Forms.Label lblPeople;
        private System.Windows.Forms.TextBox txtPeople;
        private System.Windows.Forms.Label lblMainInGAK;
        private System.Windows.Forms.TextBox txtToMain;
        private System.Windows.Forms.Label lblLect;
        private System.Windows.Forms.TextBox txtLect;
        private System.Windows.Forms.Label lblOuter;
        private System.Windows.Forms.TextBox txtOuter;
        private System.Windows.Forms.Label lblGeneralScheme;
        private System.Windows.Forms.Label lblGen;
        private System.Windows.Forms.Label lblInner;
        private System.Windows.Forms.TextBox txtInner;
        private System.Windows.Forms.Label lblList;
        private System.Windows.Forms.ComboBox cmbDiplomaList;
        private System.Windows.Forms.Label lblSubscribe;
        private System.Windows.Forms.TextBox txtSubscribe;
        private System.Windows.Forms.TextBox txtSecr;
        private System.Windows.Forms.Label lblSecr;
        private System.Windows.Forms.Label lblToOuter;
        private System.Windows.Forms.TextBox txtToOuter;
        private System.Windows.Forms.TextBox txtToInner;
        private System.Windows.Forms.Label lblToInner;
        private System.Windows.Forms.TextBox txtMain;
        private System.Windows.Forms.Label lblMain;
        private System.Windows.Forms.TextBox txtToSecretary;
        private System.Windows.Forms.Label lblToSecretary;
    }
}