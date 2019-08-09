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
    public partial class frmEditDegree : Form
    {
        public frmEditDegree()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditDegree_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы степеней
            if (mdlData.colDegree.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillDegreeList();
            }
            //при неудачной загрузке элементов из таблицы степеней
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillDegreeList()
        {
            int NumFix = 0;
            NumFix = cmbDegreeList.SelectedIndex;
            //Очищаем список
            cmbDegreeList.Items.Clear();

            //Заполняем комбо-список степенями
            for (int i = 0; i <= mdlData.colDegree.Count - 1; i++)
            {
                cmbDegreeList.Items.Add(mdlData.colDegree[i].Code + ". " + mdlData.colDegree[i].Degree);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbDegreeList.SelectedIndex = 0;
            }
            else
            {
                cmbDegreeList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colDegree.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbDegreeList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtDegree.Text = mdlData.colDegree[cmbDegreeList.SelectedIndex].Degree;
            txtShort.Text = mdlData.colDegree[cmbDegreeList.SelectedIndex].Short;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colDegree[cmbDegreeList.SelectedIndex].Degree = txtDegree.Text;
            mdlData.colDegree[cmbDegreeList.SelectedIndex].Short = txtShort.Text;
            FillDegreeList();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Степень"
            clsDegree Dee = new clsDegree();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Dee.Code = mdlData.colDegree.Count + 1;
            //Формируем задел названия для новой степени
            Dee.Degree = "Новая степень наук";
            Dee.Short = "Н.ст.н.";
            //Добавляем объект в коллекцию
            mdlData.colDegree.Add(Dee);
            //Заносим объект в список
            cmbDegreeList.Items.Add(mdlData.colDegree[mdlData.colDegree.Count - 1].Code + ". " + mdlData.colDegree[mdlData.colDegree.Count - 1].Degree);
            //Переходим к новому элементу списка
            cmbDegreeList.SelectedIndex = cmbDegreeList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbDegreeList.SelectedIndex == cmbDegreeList.Items.Count - 1)
            {
                //Проходим всех преподавателей и убираем у них
                //значение удаляемой степени
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Degree == null))
                    {
                        if (mdlData.colLecturer[i].Degree.Degree.Equals(mdlData.colDegree[cmbDegreeList.SelectedIndex].Degree))
                        {
                            mdlData.colLecturer[i].Degree = null;
                        }
                    }
                }
                //Удаляем степень из коллекции степеней
                mdlData.colDegree.RemoveAt(mdlData.colDegree.Count - 1);
                //Удаляем степень из списка степеней
                cmbDegreeList.Items.RemoveAt(cmbDegreeList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbDegreeList.SelectedIndex = cmbDegreeList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbDegreeList.SelectedIndex;
                //Проходим всех преподавателей и убираем у них
                //значение удаляемой степени
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Degree == null))
                    {
                        if (mdlData.colLecturer[i].Degree.Degree.Equals(mdlData.colDegree[DelElem].Degree))
                        {
                            mdlData.colLecturer[i].Degree = null;
                        }
                    }
                }
                
                mdlData.colDegree.RemoveAt(DelElem);
                cmbDegreeList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colDegree.Count - 1; i++)
                {
                    mdlData.colDegree[i].Code = mdlData.colDegree[i].Code - 1;
                }
                
                cmbDegreeList.SelectedIndex = DelElem;
                FillDegreeList();
            }
        }

    }
}
