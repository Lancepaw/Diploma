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
    public partial class frmEditSickList : Form
    {
        public frmEditSickList()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditSickList_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы больничных листов
            if (mdlData.colSickList.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                //Сначала загружаем список преподавателей
                FillLecturerList();
                //и список семестров
                FillSemestrList();
                //Только потом основной список больничных листов
                FillSickList();
            }
            //при неудачной загрузке элементов из таблицы больничных листов
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillSemestrList()
        {
            //Очищаем список
            cmbSemestrList.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestrList.Items.Add(mdlData.colSemestr[i].Code + ". " + 
                    mdlData.colSemestr[i].SemNum);
            }
        }

        private void FillLecturerList()
        {
            //Очищаем список
            cmbLecturerList.Items.Clear();

            //Заполняем комбо-список должностями
            //и попутно считаем суммарную ставку
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturerList.Items.Add(mdlData.colLecturer[i].Code + ". " + 
                    mdlData.colLecturer[i].FIO);
            }
        }

        private void FillSickList()
        {
            int NumFix = 0;
            NumFix = cmbLists.SelectedIndex;
            //Очищаем список
            cmbLists.Items.Clear();

            //Заполняем комбо-список номерами курсов
            for (int i = 0; i <= mdlData.colSickList.Count - 1; i++)
            {
                cmbLists.Items.Add(mdlData.colSickList[i].Code + ". " + 
                    mdlData.colSickList[i].OpenDate.Day + " " + 
                    mdlData.getMonthStringRP(mdlData.colSickList[i].OpenDate.Month) + " " +
                    mdlData.colSickList[i].OpenDate.Year + " - " + 
                    mdlData.colSickList[i].CloseDate.Day + " " +
                    mdlData.getMonthStringRP(mdlData.colSickList[i].CloseDate.Month) + " " +
                    mdlData.colSickList[i].CloseDate.Year);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbLists.SelectedIndex = 0;
            }
            else
            {
                cmbLists.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colSickList.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbLists_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mdlData.colSickList[cmbLists.SelectedIndex].Lecturer != null)
            {
                if (cmbLecturerList.Items.Count > 0)
                {
                    cmbLecturerList.SelectedIndex =
                        mdlData.colLecturer.IndexOf(mdlData.colSickList[cmbLists.SelectedIndex].Lecturer);
                }
            }
            else
            {
                cmbLecturerList.SelectedIndex = -1;
            }

            if (mdlData.colSickList[cmbLists.SelectedIndex].Lecturer != null)
            {
                if (cmbSemestrList.Items.Count > 0)
                {
                    cmbSemestrList.SelectedIndex =
                        mdlData.colSemestr.IndexOf(mdlData.colSickList[cmbLists.SelectedIndex].Semestr);
                }
            }
            else
            {
                cmbSemestrList.SelectedIndex = -1;
            }



            dtpOpen.Value = mdlData.colSickList[cmbLists.SelectedIndex].OpenDate;
            dtpClose.Value = mdlData.colSickList[cmbLists.SelectedIndex].CloseDate;

            txtDescript.Text = mdlData.colSickList[cmbLists.SelectedIndex].Descript;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            clsSickList SL = new clsSickList();
            int MaxNum = 0;
            
            for (int i = 0; i < mdlData.colSickList.Count; i++)
            {
                if (MaxNum < mdlData.colSickList[i].Code)
                {
                    MaxNum = mdlData.colSickList[i].Code;
                }
            }

            SL.Code = MaxNum + 1;
            SL.Lecturer = null;
            SL.Semestr = null;
            SL.OpenDate = DateTime.Now;
            SL.CloseDate = DateTime.Now;
            SL.Descript = "Новое примечание";

            //Заносим элемент в коллекцию
            mdlData.colSickList.Add(SL);

            //Заносим объект в список
            cmbLists.Items.Add(SL.Code + ". " +
                    SL.OpenDate.Day + " " +
                    mdlData.getMonthStringRP(SL.OpenDate.Month) + " " +
                    SL.OpenDate.Year + " - " +
                    SL.CloseDate.Day + " " +
                    mdlData.getMonthStringRP(SL.CloseDate.Month) + " " +
                    SL.CloseDate.Year);
            //Переходим к новому элементу списка
            cmbLists.SelectedIndex = cmbLists.Items.Count - 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (cmbLecturerList.SelectedIndex >= 0)
            {
                mdlData.colSickList[cmbLists.SelectedIndex].Lecturer = mdlData.colLecturer[cmbLecturerList.SelectedIndex];
            }
            else
            {
                mdlData.colSickList[cmbLists.SelectedIndex].Lecturer = null;
            }

            if (cmbSemestrList.SelectedIndex >= 0)
            {
                mdlData.colSickList[cmbLists.SelectedIndex].Semestr = mdlData.colSemestr[cmbSemestrList.SelectedIndex];
            }
            else
            {
                mdlData.colSickList[cmbLists.SelectedIndex].Semestr = null;
            }

            mdlData.colSickList[cmbLists.SelectedIndex].OpenDate = dtpOpen.Value;
            mdlData.colSickList[cmbLists.SelectedIndex].CloseDate = dtpClose.Value;
            mdlData.colSickList[cmbLists.SelectedIndex].Descript = txtDescript.Text;

            FillSickList();
        }
    }
}
