using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataBaseIO
{
    public partial class frmDopWork : Form
    {

        public int DWCode;
        public int SemIndex;
        public int LectIndex;
        public bool flgLocalNeedSaveWork;
        public bool flgLocalNeedSaveComment;
        public bool flgLocalNeedSaveVolume;
        public bool flgLocalNeedSaveDate;

        public bool flgReturnedState = false;
        public bool flgLoadText = false;

        public Color EnabledState;
        public Color DisabledState;

        public frmDopWork()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmDopWork_Load(object sender, EventArgs e)
        {
            optNIR.Checked = false;
            optUMR.Checked = false;
            optOMR.Checked = false;

            EnabledState = txtWork.BackColor;
            DisabledState = Color.LightGray;
            
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы доп.работы
            if (mdlData.colDopWork.Count > 0)
            {
                optNIR.Enabled = true;
                optUMR.Enabled = true;
                optOMR.Enabled = true;

                FillSemestrList();
                FillLecturerList();
            }
            //при неудачной загрузке элементов из таблицы доп.работы
            else
            {
                optNIR.Enabled = false;
                optUMR.Enabled = false;
                optOMR.Enabled = false;
            }

            flgLocalNeedSaveWork = false;
            flgLocalNeedSaveComment = false;
            flgLocalNeedSaveVolume = false;
            flgLocalNeedSaveDate = false;
        }

        private void FillLecturerList()
        {
            int NumFix = 0;
            NumFix = cmbLectList.SelectedIndex;
            //Очищаем список
            cmbLectList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLectList.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbLectList.SelectedIndex = 0;
            }
            else
            {
                cmbLectList.SelectedIndex = NumFix;
            }
        }

        private void FillSemestrList()
        {
            int NumFix = 0;
            NumFix = cmbSemList.SelectedIndex;
            //Очищаем список
            cmbSemList.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemList.Items.Add(mdlData.colSemestr[i].Code + ". " + mdlData.colSemestr[i].SemNum);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSemList.SelectedIndex = 0;
            }
            else
            {
                cmbSemList.SelectedIndex = NumFix;
            }
        }

        //Изменение состояния опции на НИР
        private void optNIR_CheckedChanged(object sender, EventArgs e)
        {
            //int NumFix = 0;
            bool flgExist;

            //Если ни один признак необходимости сохранения не выставлен,
            //то можно поменять состояние окон отображения текста
            if (!flgLocalNeedSaveWork & !flgLocalNeedSaveComment & !flgLocalNeedSaveVolume & !flgLocalNeedSaveDate)
            {
                flgLoadText = true;
                txtWork.Text = "";

                flgExist = false;
                for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
                {
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemList.SelectedIndex].SemNum))
                    {
                        if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLectList.SelectedIndex].FIO))
                        {
                            txtWork.Text = mdlData.colDopWork[i].NIR;
                            txtDate.Text = mdlData.colDopWork[i].DateNIR;
                            txtVolume.Text = mdlData.colDopWork[i].VolumeNIR;
                            txtComment.Text = mdlData.colDopWork[i].CommNIR;
                            DWCode = i;
                            flgLoadText = false;
                            flgExist = true;
                        }
                    }
                }

                if (!flgExist)
                {
                    txtWork.Text = "(запись не найдена)";
                }

                flgLoadText = false;
            }
            else
            {
                MessageMe();
            }
        }

        private void MessageMe()
        {
            if (flgLocalNeedSaveWork)
            {
                MessageBox.Show("Наименование работы не сохранено");
                return;
            }

            if (flgLocalNeedSaveComment)
            {
                MessageBox.Show("Примечание не сохранено");
                return;
            }

            if (flgLocalNeedSaveVolume)
            {
                MessageBox.Show("Объём не сохранён");
                return;
            }
        }

        private void optUMR_CheckedChanged(object sender, EventArgs e)
        {
            //int NumFix = 0;
            bool flgExist;

            //Если ни один признак необходимости сохранения не выставлен,
            //то можно поменять состояние окон отображения текста
            if (!flgLocalNeedSaveWork & !flgLocalNeedSaveComment & !flgLocalNeedSaveVolume & !flgLocalNeedSaveDate)
            {
                flgLoadText = true;
                txtWork.Text = "";

                flgExist = false;
                for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
                {
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemList.SelectedIndex].SemNum))
                    {
                        if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLectList.SelectedIndex].FIO))
                        {
                            txtWork.Text = mdlData.colDopWork[i].UMR;
                            txtDate.Text = mdlData.colDopWork[i].DateUMR;
                            txtVolume.Text = mdlData.colDopWork[i].VolumeUMR;
                            txtComment.Text = mdlData.colDopWork[i].CommUMR;
                            DWCode = i;
                            flgLoadText = false;
                            flgExist = true;
                        }
                    }
                }

                if (!flgExist)
                {
                    txtWork.Text = "(запись не найдена)";
                }

                flgLoadText = false;
            }
            else
            {
                MessageMe();
            }
        }

        private void optOMR_CheckedChanged(object sender, EventArgs e)
        {
            //int NumFix = 0;
            bool flgExist;

            //Если ни один признак необходимости сохранения не выставлен,
            //то можно поменять состояние окон отображения текста
            if (!flgLocalNeedSaveWork & !flgLocalNeedSaveComment & !flgLocalNeedSaveVolume & !flgLocalNeedSaveDate)
            {
                flgLoadText = true;
                txtWork.Text = "";

                flgExist = false;
                for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
                {
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemList.SelectedIndex].SemNum))
                    {
                        if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLectList.SelectedIndex].FIO))
                        {
                            txtWork.Text = mdlData.colDopWork[i].OMR;
                            txtDate.Text = mdlData.colDopWork[i].DateOMR;
                            txtVolume.Text = mdlData.colDopWork[i].VolumeOMR;
                            txtComment.Text = mdlData.colDopWork[i].CommOMR;
                            DWCode = i;
                            flgLoadText = false;
                            flgExist = true;
                        }
                    }
                }

                if (!flgExist)
                {
                    txtWork.Text = "(запись не найдена)";
                }

                flgLoadText = false;
            }
            else
            {
                MessageMe();
            }
        }

        //Изменение индекса в списке семестров влечёт за собой следующие действия
        private void cmbSemList_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Если переход сюда состоялся вручную, то выполняется следующая группа действий
            if (!flgReturnedState)
            {
                //Если ничего не менялось, то свободно переходить к другому индексу
                if (!flgLocalNeedSaveWork & !flgLocalNeedSaveComment & !flgLocalNeedSaveVolume & !flgLocalNeedSaveDate)
                {
                    SemIndex = cmbSemList.SelectedIndex;
                    ResetState();
                }
                else
                {
                    //Если пользователь подтверждает смену индекса без сохранения изменений
                    if (MessageBox.Show("Не сохранено. Продолжить?",
                        "Сохранение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        //Сохраняем новый индекс, поменянный пользователе
                        SemIndex = cmbSemList.SelectedIndex;
                        //сбрасываем все выставленные признаки необходимости сохранения
                        ResetState();
                    }
                    else
                    {
                        //выставляем признак возврата предыдущего состояния
                        flgReturnedState = true;
                        //Меняем индекс на предыдущий, сохранённый
                        cmbSemList.SelectedIndex = SemIndex;
                    }
                }
            }
            //Если переход сюда случился автоматом в связи с возвратом
            //предыдущего состояния списка
            else
            {
                //сбрасываем признак возврата предыдущего состояния
                flgReturnedState = false;
            }
        }

        private void cmbLectList_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Если переход сюда состоялся вручную, то выполняется следующая группа действий
            if (!flgReturnedState)
            {
                //Если ничего не менялось, то свободно переходить к другому индексу
                if (!flgLocalNeedSaveWork & !flgLocalNeedSaveComment & !flgLocalNeedSaveVolume & !flgLocalNeedSaveDate)
                {
                    LectIndex = cmbLectList.SelectedIndex;
                    ResetState();
                }
                else
                {
                    if (MessageBox.Show("Не сохранено. Продолжить?",
                        "Сохранение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        LectIndex = cmbLectList.SelectedIndex;
                        ResetState();
                    }
                    else
                    {
                        //выставляем признак возврата предыдущего состояния
                        flgReturnedState = true;
                        cmbLectList.SelectedIndex = LectIndex;
                    }
                }
            }
            //Если переход сюда случился автоматом в связи с возвратом
            //предыдущего состояния списка
            else
            {
                //сбрасываем признак возврата предыдущего состояния
                flgReturnedState = false;
            }
        }

        private void ResetState()
        {
            bool NIR_prev;
            bool UMR_prev;
            bool OMR_prev;

            flgLocalNeedSaveWork = false;
            txtWork.BackColor = EnabledState;
            flgLocalNeedSaveComment = false;
            txtComment.BackColor = EnabledState;
            flgLocalNeedSaveVolume = false;
            txtVolume.BackColor = EnabledState;
            flgLocalNeedSaveDate = false;
            txtDate.BackColor = EnabledState;

            NIR_prev = optNIR.Checked;
            UMR_prev = optUMR.Checked;
            OMR_prev = optOMR.Checked;

            optNIR.Checked = false;
            optUMR.Checked = false;
            optOMR.Checked = false;

            optNIR.Checked = NIR_prev;
            optUMR.Checked = UMR_prev;
            optOMR.Checked = OMR_prev;
        }
        
        //Сохранение
        private void btnSave_Click(object sender, EventArgs e)
        {
            //Если отмечена учебно-методическая работа
            if (optUMR.Checked)
            {
                mdlData.colDopWork[DWCode].UMR = txtWork.Text;
                mdlData.colDopWork[DWCode].DateUMR = txtDate.Text;
                mdlData.colDopWork[DWCode].VolumeUMR = txtVolume.Text;
                mdlData.colDopWork[DWCode].CommUMR = txtComment.Text;
            }
            else
            {
                //Если отмечена научно-исследовательская работа
                if (optNIR.Checked)
                {
                    mdlData.colDopWork[DWCode].NIR = txtWork.Text;
                    mdlData.colDopWork[DWCode].DateNIR = txtDate.Text;
                    mdlData.colDopWork[DWCode].VolumeNIR = txtVolume.Text;
                    mdlData.colDopWork[DWCode].CommNIR = txtComment.Text;
                }
                else
                {
                    //Если отмечена организационно-методическая работа
                    if (optOMR.Checked)
                    {
                        mdlData.colDopWork[DWCode].OMR = txtWork.Text;
                        mdlData.colDopWork[DWCode].DateOMR = txtDate.Text;
                        mdlData.colDopWork[DWCode].VolumeOMR = txtVolume.Text;
                        mdlData.colDopWork[DWCode].CommOMR = txtComment.Text;
                    }
                }
            }

            flgLocalNeedSaveWork = false;
            txtWork.BackColor = EnabledState;
            flgLocalNeedSaveComment = false;
            txtComment.BackColor = EnabledState;
            flgLocalNeedSaveVolume = false;
            txtVolume.BackColor = EnabledState;
            flgLocalNeedSaveDate = false;
            txtDate.BackColor = EnabledState;
        }

        //Если текст работы поменялся
        private void txtWork_TextChanged(object sender, EventArgs e)
        {
            if (!flgLoadText)
            {
                //Выставляем флаг произошедших изменений
                flgLocalNeedSaveWork = true;
                //Перекрашиваем подложку компонента
                txtWork.BackColor = DisabledState;
            }
        }

        private void txtComment_TextChanged(object sender, EventArgs e)
        {
            if (!flgLoadText)
            {
                flgLocalNeedSaveComment = true;
                txtComment.BackColor = DisabledState;
            }
        }

        private void txtVolume_TextChanged(object sender, EventArgs e)
        {
            if (!flgLoadText)
            {
                flgLocalNeedSaveVolume = true;
                txtVolume.BackColor = DisabledState;
            }
        }

        private void txtDate_TextChanged(object sender, EventArgs e)
        {
            if (!flgLoadText)
            {
                flgLocalNeedSaveDate = true;
                txtDate.BackColor = DisabledState;
            }
        }
    }
}
