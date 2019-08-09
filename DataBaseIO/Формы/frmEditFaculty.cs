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
    public partial class frmEditFaculty : Form
    {
        public frmEditFaculty()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colFaculty[cmbFacultyList.SelectedIndex].Faculty = txtFaculty.Text;
            mdlData.colFaculty[cmbFacultyList.SelectedIndex].Short = txtShort.Text;
            mdlData.colFaculty[cmbFacultyList.SelectedIndex].Diff = txtDiff.Text;
        }

        private void cmbFacultyList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtFaculty.Text = mdlData.colFaculty[cmbFacultyList.SelectedIndex].Faculty;
            txtShort.Text = mdlData.colFaculty[cmbFacultyList.SelectedIndex].Short;
            txtDiff.Text = mdlData.colFaculty[cmbFacultyList.SelectedIndex].Diff;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Факультет"
            clsFaculty Fac = new clsFaculty();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Fac.Code = mdlData.colFaculty.Count + 1;
            //Формируем задел названия для нового факультета
            Fac.Faculty = "Новый факультет";
            Fac.Short = "Нов.фак.";
            Fac.Diff = "Б";
            //Добавляем объект в коллекцию
            mdlData.colFaculty.Add(Fac);
            //Заносим объект в список
            cmbFacultyList.Items.Add(mdlData.colFaculty[mdlData.colFaculty.Count - 1].Code + ". " + mdlData.colFaculty[mdlData.colFaculty.Count - 1].Faculty);
            //Переходим к новому элементу списка
            cmbFacultyList.SelectedIndex = cmbFacultyList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbFacultyList.SelectedIndex == cmbFacultyList.Items.Count - 1)
            {
                //Проходим все специальности и убираем у них
                //значение удаляемого факультета
                for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
                {
                    if (!(mdlData.colSpecialisation[i].Faculty == null))
                    {
                        if (mdlData.colSpecialisation[i].Faculty.Short.Equals(mdlData.colFaculty[cmbFacultyList.SelectedIndex].Short))
                        {
                            mdlData.colSpecialisation[i].Faculty = null;
                        }
                    }
                }
                //Удаляем факультет из коллекции факультетов
                mdlData.colFaculty.RemoveAt(mdlData.colFaculty.Count - 1);
                //Удаляем факультет из списка факультетов
                cmbFacultyList.Items.RemoveAt(cmbFacultyList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbFacultyList.SelectedIndex = cmbFacultyList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbFacultyList.SelectedIndex;
                //Проходим все специальности и убираем у них
                //значение удаляемого факультета
                for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
                {
                    if (!(mdlData.colSpecialisation[i].Faculty == null))
                    {
                        if (mdlData.colSpecialisation[i].Faculty.Short.Equals(mdlData.colFaculty[DelElem].Short))
                        {
                            mdlData.colSpecialisation[i].Faculty = null;
                        }
                    }
                }

                mdlData.colFaculty.RemoveAt(DelElem);
                cmbFacultyList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colFaculty.Count - 1; i++)
                {
                    mdlData.colFaculty[i].Code = mdlData.colFaculty[i].Code - 1;
                }

                cmbFacultyList.SelectedIndex = DelElem;
                FillFacultyList();
            }
        }

        private void frmEditFaculty_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы факультетов
            if (mdlData.colFaculty.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillFacultyList();
            }
            //при неудачной загрузке элементов из таблицы факультетов
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillFacultyList()
        {
            int NumFix = 0;
            NumFix = cmbFacultyList.SelectedIndex;
            //Очищаем список
            cmbFacultyList.Items.Clear();

            //Заполняем комбо-список степенями
            for (int i = 0; i <= mdlData.colFaculty.Count - 1; i++)
            {
                cmbFacultyList.Items.Add(mdlData.colFaculty[i].Code + ". " + mdlData.colFaculty[i].Short);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbFacultyList.SelectedIndex = 0;
            }
            else
            {
                cmbFacultyList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colFaculty.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }
    }
}
