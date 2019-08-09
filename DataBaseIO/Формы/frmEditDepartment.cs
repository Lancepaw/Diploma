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
    public partial class frmEditDepartment : Form
    {
        public frmEditDepartment()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditDepartment_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы кафедр
            if (mdlData.colDepart.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillDepartmentList();
            }
            //при неудачной загрузке элементов из таблицы кафедр
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillDepartmentList()
        {
            int NumFix = 0;
            NumFix = cmbDepartList.SelectedIndex;
            //Очищаем список
            cmbDepartList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                cmbDepartList.Items.Add(mdlData.colDepart[i].Code + ". " + mdlData.colDepart[i].Short);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbDepartList.SelectedIndex = 0;
            }
            else
            {
                cmbDepartList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colDepart.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbDepartList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtName.Text = mdlData.colDepart[cmbDepartList.SelectedIndex].Kafedra;
            txtShort.Text = mdlData.colDepart[cmbDepartList.SelectedIndex].Short;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colDepart[cmbDepartList.SelectedIndex].Kafedra = txtName.Text;
            mdlData.colDepart[cmbDepartList.SelectedIndex].Short = txtShort.Text;
            FillDepartmentList();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Кафедра"
            clsDepartment Dp = new clsDepartment();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Dp.Code = mdlData.colDepart.Count + 1;
            //Формируем задел названия для новой кафедры
            Dp.Kafedra = "Новая кафедра";
            Dp.Short = "Нов. каф.";
            //Добавляем объект в коллекцию
            mdlData.colDepart.Add(Dp);
            //Заносим объект в список
            cmbDepartList.Items.Add(mdlData.colDepart[mdlData.colDepart.Count - 1].Code + ". " + mdlData.colDepart[mdlData.colDepart.Count - 1].Kafedra);
            //Переходим к новому элементу списка
            cmbDepartList.SelectedIndex = cmbDepartList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbDepartList.SelectedIndex == cmbDepartList.Items.Count - 1)
            {
                //Проходим всех преподавателей и убираем у них
                //значение удаляемой кафедры
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Depart == null))
                    {
                        if (mdlData.colLecturer[i].Depart.Kafedra.Equals(mdlData.colDepart[cmbDepartList.SelectedIndex].Kafedra))
                        {
                            mdlData.colLecturer[i].Depart = null;
                        }
                    }
                }
                mdlData.colDepart.RemoveAt(mdlData.colDepart.Count - 1);
                cmbDepartList.Items.RemoveAt(cmbDepartList.Items.Count - 1);
                cmbDepartList.SelectedIndex = cmbDepartList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                DelElem = cmbDepartList.SelectedIndex;

                //Проходим всех преподавателей и убираем у них
                //значение удаляемой кафедры
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Depart == null))
                    {
                        if (mdlData.colLecturer[i].Depart.Kafedra.Equals(mdlData.colDepart[DelElem].Kafedra))
                        {
                            mdlData.colLecturer[i].Depart = null;
                        }
                    }
                }

                mdlData.colDepart.RemoveAt(DelElem);
                cmbDepartList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colDepart.Count - 1; i++)
                {
                    mdlData.colDepart[i].Code = mdlData.colDepart[i].Code - 1;
                }
                cmbDepartList.SelectedIndex = DelElem;
                FillDepartmentList();
            }
        }
    }
}
