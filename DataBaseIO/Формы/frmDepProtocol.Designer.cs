namespace DataBaseIO
{
    partial class frmDepProtocol
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
            this.chkLstDepWorkers = new System.Windows.Forms.CheckedListBox();
            this.lblLecturerList = new System.Windows.Forms.Label();
            this.btnWord = new System.Windows.Forms.Button();
            this.lblProtocolNumList = new System.Windows.Forms.Label();
            this.cmbProtocolNumList = new System.Windows.Forms.ComboBox();
            this.cldrMain = new System.Windows.Forms.MonthCalendar();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnAnnounce = new System.Windows.Forms.Button();
            this.lblTime = new System.Windows.Forms.Label();
            this.txtTime = new System.Windows.Forms.TextBox();
            this.lblRoom = new System.Windows.Forms.Label();
            this.txtRoom = new System.Windows.Forms.TextBox();
            this.btnPlan = new System.Windows.Forms.Button();
            this.lstQuestions = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.btnPaste = new System.Windows.Forms.Button();
            this.btnCopyAll = new System.Windows.Forms.Button();
            this.lblSpeaker1 = new System.Windows.Forms.Label();
            this.lblSpeaker2 = new System.Windows.Forms.Label();
            this.lblSpeaker3 = new System.Windows.Forms.Label();
            this.txtQuestion = new System.Windows.Forms.TextBox();
            this.lblQuestion = new System.Windows.Forms.Label();
            this.cmbSpeaker1 = new System.Windows.Forms.ComboBox();
            this.cmbSpeaker2 = new System.Windows.Forms.ComboBox();
            this.cmbSpeaker3 = new System.Windows.Forms.ComboBox();
            this.cmbSpeaker4 = new System.Windows.Forms.ComboBox();
            this.cmbSpeaker5 = new System.Windows.Forms.ComboBox();
            this.lblSpeaker4 = new System.Windows.Forms.Label();
            this.lblSpeaker5 = new System.Windows.Forms.Label();
            this.btnCut = new System.Windows.Forms.Button();
            this.btnChange = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(710, 416);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(62, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // chkLstDepWorkers
            // 
            this.chkLstDepWorkers.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.chkLstDepWorkers.CheckOnClick = true;
            this.chkLstDepWorkers.FormattingEnabled = true;
            this.chkLstDepWorkers.Location = new System.Drawing.Point(12, 34);
            this.chkLstDepWorkers.Name = "chkLstDepWorkers";
            this.chkLstDepWorkers.Size = new System.Drawing.Size(247, 379);
            this.chkLstDepWorkers.TabIndex = 1;
            // 
            // lblLecturerList
            // 
            this.lblLecturerList.AutoSize = true;
            this.lblLecturerList.Location = new System.Drawing.Point(12, 9);
            this.lblLecturerList.Name = "lblLecturerList";
            this.lblLecturerList.Size = new System.Drawing.Size(143, 13);
            this.lblLecturerList.TabIndex = 2;
            this.lblLecturerList.Text = "Отметить присутствующих";
            // 
            // btnWord
            // 
            this.btnWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnWord.Location = new System.Drawing.Point(498, 416);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(83, 23);
            this.btnWord.TabIndex = 3;
            this.btnWord.Text = "Протокол";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // lblProtocolNumList
            // 
            this.lblProtocolNumList.AutoSize = true;
            this.lblProtocolNumList.Location = new System.Drawing.Point(268, 205);
            this.lblProtocolNumList.Name = "lblProtocolNumList";
            this.lblProtocolNumList.Size = new System.Drawing.Size(100, 13);
            this.lblProtocolNumList.TabIndex = 4;
            this.lblProtocolNumList.Text = "Номер протокола:";
            // 
            // cmbProtocolNumList
            // 
            this.cmbProtocolNumList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbProtocolNumList.FormattingEnabled = true;
            this.cmbProtocolNumList.Location = new System.Drawing.Point(374, 202);
            this.cmbProtocolNumList.Name = "cmbProtocolNumList";
            this.cmbProtocolNumList.Size = new System.Drawing.Size(61, 21);
            this.cmbProtocolNumList.TabIndex = 5;
            this.cmbProtocolNumList.SelectedIndexChanged += new System.EventHandler(this.cmbProtocolNumList_SelectedIndexChanged);
            // 
            // cldrMain
            // 
            this.cldrMain.Location = new System.Drawing.Point(271, 34);
            this.cldrMain.Name = "cldrMain";
            this.cldrMain.TabIndex = 6;
            this.cldrMain.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.cldrMain_DateChanged);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(12, 416);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(89, 23);
            this.btnSelectAll.TabIndex = 7;
            this.btnSelectAll.Text = "Выбрать всех";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(184, 416);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 8;
            this.btnClear.Text = "Очистить";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnAnnounce
            // 
            this.btnAnnounce.Location = new System.Drawing.Point(640, 416);
            this.btnAnnounce.Name = "btnAnnounce";
            this.btnAnnounce.Size = new System.Drawing.Size(64, 23);
            this.btnAnnounce.TabIndex = 9;
            this.btnAnnounce.Text = "Повестка";
            this.btnAnnounce.UseVisualStyleBackColor = true;
            this.btnAnnounce.Click += new System.EventHandler(this.btnAnnounce_Click);
            // 
            // lblTime
            // 
            this.lblTime.AutoSize = true;
            this.lblTime.Location = new System.Drawing.Point(268, 232);
            this.lblTime.Name = "lblTime";
            this.lblTime.Size = new System.Drawing.Size(43, 13);
            this.lblTime.TabIndex = 10;
            this.lblTime.Text = "Время:";
            // 
            // txtTime
            // 
            this.txtTime.Location = new System.Drawing.Point(374, 229);
            this.txtTime.Name = "txtTime";
            this.txtTime.Size = new System.Drawing.Size(61, 20);
            this.txtTime.TabIndex = 11;
            this.txtTime.Text = "15:10";
            // 
            // lblRoom
            // 
            this.lblRoom.AutoSize = true;
            this.lblRoom.Location = new System.Drawing.Point(268, 260);
            this.lblRoom.Name = "lblRoom";
            this.lblRoom.Size = new System.Drawing.Size(63, 13);
            this.lblRoom.TabIndex = 12;
            this.lblRoom.Text = "Аудитория:";
            // 
            // txtRoom
            // 
            this.txtRoom.Location = new System.Drawing.Point(374, 255);
            this.txtRoom.Name = "txtRoom";
            this.txtRoom.Size = new System.Drawing.Size(61, 20);
            this.txtRoom.TabIndex = 13;
            this.txtRoom.Text = "4527";
            // 
            // btnPlan
            // 
            this.btnPlan.Location = new System.Drawing.Point(587, 416);
            this.btnPlan.Name = "btnPlan";
            this.btnPlan.Size = new System.Drawing.Size(47, 23);
            this.btnPlan.TabIndex = 14;
            this.btnPlan.Text = "План";
            this.btnPlan.UseVisualStyleBackColor = true;
            // 
            // lstQuestions
            // 
            this.lstQuestions.FormattingEnabled = true;
            this.lstQuestions.HorizontalScrollbar = true;
            this.lstQuestions.Location = new System.Drawing.Point(447, 34);
            this.lstQuestions.Name = "lstQuestions";
            this.lstQuestions.Size = new System.Drawing.Size(325, 160);
            this.lstQuestions.TabIndex = 15;
            this.lstQuestions.SelectedIndexChanged += new System.EventHandler(this.lstQuestions_SelectedIndexChanged);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(447, 387);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(104, 23);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(557, 387);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(104, 23);
            this.btnDel.TabIndex = 17;
            this.btnDel.Text = "Удалить";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(667, 387);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(104, 23);
            this.btnClearAll.TabIndex = 18;
            this.btnClearAll.Text = "Очистить";
            this.btnClearAll.UseVisualStyleBackColor = true;
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(667, 250);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(104, 23);
            this.btnCopy.TabIndex = 19;
            this.btnCopy.Text = "Скопировать";
            this.btnCopy.UseVisualStyleBackColor = true;
            // 
            // btnPaste
            // 
            this.btnPaste.Location = new System.Drawing.Point(667, 358);
            this.btnPaste.Name = "btnPaste";
            this.btnPaste.Size = new System.Drawing.Size(104, 23);
            this.btnPaste.TabIndex = 20;
            this.btnPaste.Text = "Вставить";
            this.btnPaste.UseVisualStyleBackColor = true;
            // 
            // btnCopyAll
            // 
            this.btnCopyAll.Location = new System.Drawing.Point(667, 277);
            this.btnCopyAll.Name = "btnCopyAll";
            this.btnCopyAll.Size = new System.Drawing.Size(104, 23);
            this.btnCopyAll.TabIndex = 21;
            this.btnCopyAll.Text = "Скопировать всё";
            this.btnCopyAll.UseVisualStyleBackColor = true;
            // 
            // lblSpeaker1
            // 
            this.lblSpeaker1.AutoSize = true;
            this.lblSpeaker1.Location = new System.Drawing.Point(444, 255);
            this.lblSpeaker1.Name = "lblSpeaker1";
            this.lblSpeaker1.Size = new System.Drawing.Size(72, 13);
            this.lblSpeaker1.TabIndex = 22;
            this.lblSpeaker1.Text = "Докладчик 1";
            // 
            // lblSpeaker2
            // 
            this.lblSpeaker2.AutoSize = true;
            this.lblSpeaker2.Location = new System.Drawing.Point(444, 282);
            this.lblSpeaker2.Name = "lblSpeaker2";
            this.lblSpeaker2.Size = new System.Drawing.Size(72, 13);
            this.lblSpeaker2.TabIndex = 23;
            this.lblSpeaker2.Text = "Докладчик 2";
            // 
            // lblSpeaker3
            // 
            this.lblSpeaker3.AutoSize = true;
            this.lblSpeaker3.Location = new System.Drawing.Point(444, 309);
            this.lblSpeaker3.Name = "lblSpeaker3";
            this.lblSpeaker3.Size = new System.Drawing.Size(72, 13);
            this.lblSpeaker3.TabIndex = 24;
            this.lblSpeaker3.Text = "Докладчик 3";
            this.lblSpeaker3.Click += new System.EventHandler(this.lblSpeaker3_Click);
            // 
            // txtQuestion
            // 
            this.txtQuestion.Location = new System.Drawing.Point(497, 200);
            this.txtQuestion.Multiline = true;
            this.txtQuestion.Name = "txtQuestion";
            this.txtQuestion.Size = new System.Drawing.Size(275, 43);
            this.txtQuestion.TabIndex = 25;
            // 
            // lblQuestion
            // 
            this.lblQuestion.AutoSize = true;
            this.lblQuestion.Location = new System.Drawing.Point(444, 202);
            this.lblQuestion.Name = "lblQuestion";
            this.lblQuestion.Size = new System.Drawing.Size(47, 13);
            this.lblQuestion.TabIndex = 26;
            this.lblQuestion.Text = "Вопрос:";
            // 
            // cmbSpeaker1
            // 
            this.cmbSpeaker1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeaker1.FormattingEnabled = true;
            this.cmbSpeaker1.Location = new System.Drawing.Point(522, 252);
            this.cmbSpeaker1.Name = "cmbSpeaker1";
            this.cmbSpeaker1.Size = new System.Drawing.Size(139, 21);
            this.cmbSpeaker1.TabIndex = 27;
            // 
            // cmbSpeaker2
            // 
            this.cmbSpeaker2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeaker2.FormattingEnabled = true;
            this.cmbSpeaker2.Location = new System.Drawing.Point(522, 279);
            this.cmbSpeaker2.Name = "cmbSpeaker2";
            this.cmbSpeaker2.Size = new System.Drawing.Size(139, 21);
            this.cmbSpeaker2.TabIndex = 28;
            // 
            // cmbSpeaker3
            // 
            this.cmbSpeaker3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeaker3.FormattingEnabled = true;
            this.cmbSpeaker3.Location = new System.Drawing.Point(522, 306);
            this.cmbSpeaker3.Name = "cmbSpeaker3";
            this.cmbSpeaker3.Size = new System.Drawing.Size(139, 21);
            this.cmbSpeaker3.TabIndex = 29;
            // 
            // cmbSpeaker4
            // 
            this.cmbSpeaker4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeaker4.FormattingEnabled = true;
            this.cmbSpeaker4.Location = new System.Drawing.Point(522, 333);
            this.cmbSpeaker4.Name = "cmbSpeaker4";
            this.cmbSpeaker4.Size = new System.Drawing.Size(139, 21);
            this.cmbSpeaker4.TabIndex = 30;
            // 
            // cmbSpeaker5
            // 
            this.cmbSpeaker5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeaker5.FormattingEnabled = true;
            this.cmbSpeaker5.Location = new System.Drawing.Point(522, 360);
            this.cmbSpeaker5.Name = "cmbSpeaker5";
            this.cmbSpeaker5.Size = new System.Drawing.Size(139, 21);
            this.cmbSpeaker5.TabIndex = 31;
            // 
            // lblSpeaker4
            // 
            this.lblSpeaker4.AutoSize = true;
            this.lblSpeaker4.Location = new System.Drawing.Point(444, 336);
            this.lblSpeaker4.Name = "lblSpeaker4";
            this.lblSpeaker4.Size = new System.Drawing.Size(72, 13);
            this.lblSpeaker4.TabIndex = 32;
            this.lblSpeaker4.Text = "Докладчик 4";
            // 
            // lblSpeaker5
            // 
            this.lblSpeaker5.AutoSize = true;
            this.lblSpeaker5.Location = new System.Drawing.Point(444, 363);
            this.lblSpeaker5.Name = "lblSpeaker5";
            this.lblSpeaker5.Size = new System.Drawing.Size(72, 13);
            this.lblSpeaker5.TabIndex = 33;
            this.lblSpeaker5.Text = "Докладчик 5";
            // 
            // btnCut
            // 
            this.btnCut.Location = new System.Drawing.Point(667, 331);
            this.btnCut.Name = "btnCut";
            this.btnCut.Size = new System.Drawing.Size(104, 23);
            this.btnCut.TabIndex = 34;
            this.btnCut.Text = "Вырезать";
            this.btnCut.UseVisualStyleBackColor = true;
            // 
            // btnChange
            // 
            this.btnChange.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnChange.Location = new System.Drawing.Point(667, 304);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(104, 23);
            this.btnChange.TabIndex = 35;
            this.btnChange.Text = "Изменить";
            this.btnChange.UseVisualStyleBackColor = true;
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // frmDepProtocol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 447);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.btnCut);
            this.Controls.Add(this.lblSpeaker5);
            this.Controls.Add(this.lblSpeaker4);
            this.Controls.Add(this.cmbSpeaker5);
            this.Controls.Add(this.cmbSpeaker4);
            this.Controls.Add(this.cmbSpeaker3);
            this.Controls.Add(this.cmbSpeaker2);
            this.Controls.Add(this.cmbSpeaker1);
            this.Controls.Add(this.lblQuestion);
            this.Controls.Add(this.txtQuestion);
            this.Controls.Add(this.lblSpeaker3);
            this.Controls.Add(this.lblSpeaker2);
            this.Controls.Add(this.lblSpeaker1);
            this.Controls.Add(this.btnCopyAll);
            this.Controls.Add(this.btnPaste);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnClearAll);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lstQuestions);
            this.Controls.Add(this.btnPlan);
            this.Controls.Add(this.txtRoom);
            this.Controls.Add(this.lblRoom);
            this.Controls.Add(this.txtTime);
            this.Controls.Add(this.lblTime);
            this.Controls.Add(this.btnAnnounce);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this.cldrMain);
            this.Controls.Add(this.cmbProtocolNumList);
            this.Controls.Add(this.lblProtocolNumList);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.lblLecturerList);
            this.Controls.Add(this.chkLstDepWorkers);
            this.Controls.Add(this.btnClose);
            this.Name = "frmDepProtocol";
            this.Text = "Формирование протокола кафедры";
            this.Load += new System.EventHandler(this.frmDepProtocol_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckedListBox chkLstDepWorkers;
        private System.Windows.Forms.Label lblLecturerList;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Label lblProtocolNumList;
        private System.Windows.Forms.ComboBox cmbProtocolNumList;
        private System.Windows.Forms.MonthCalendar cldrMain;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnAnnounce;
        private System.Windows.Forms.Label lblTime;
        private System.Windows.Forms.TextBox txtTime;
        private System.Windows.Forms.Label lblRoom;
        private System.Windows.Forms.TextBox txtRoom;
        private System.Windows.Forms.Button btnPlan;
        private System.Windows.Forms.ListBox lstQuestions;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Button btnPaste;
        private System.Windows.Forms.Button btnCopyAll;
        private System.Windows.Forms.Label lblSpeaker1;
        private System.Windows.Forms.Label lblSpeaker2;
        private System.Windows.Forms.Label lblSpeaker3;
        private System.Windows.Forms.TextBox txtQuestion;
        private System.Windows.Forms.Label lblQuestion;
        private System.Windows.Forms.ComboBox cmbSpeaker1;
        private System.Windows.Forms.ComboBox cmbSpeaker2;
        private System.Windows.Forms.ComboBox cmbSpeaker3;
        private System.Windows.Forms.ComboBox cmbSpeaker4;
        private System.Windows.Forms.ComboBox cmbSpeaker5;
        private System.Windows.Forms.Label lblSpeaker4;
        private System.Windows.Forms.Label lblSpeaker5;
        private System.Windows.Forms.Button btnCut;
        private System.Windows.Forms.Button btnChange;
    }
}