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
    public partial class frmEditStatus : Form
    {
        public frmEditStatus()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colStatus[cmbStatusList.SelectedIndex].Status = txtStatus.Text;
            FillStatusList();
        }

        private void frmEditStatus_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы званий
            if (mdlData.colStatus.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillStatusList();
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillStatusList()
        {
            int NumFix = 0;
            NumFix = cmbStatusList.SelectedIndex;
            //Очищаем список
            cmbStatusList.Items.Clear();

            //Заполняем комбо-список званиями
            for (int i = 0; i <= mdlData.colStatus.Count - 1; i++)
            {
                cmbStatusList.Items.Add(mdlData.colStatus[i].Code + ". " + mdlData.colStatus[i].Status);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbStatusList.SelectedIndex = 0;
            }
            else
            {
                cmbStatusList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colStatus.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbStatusList.SelectedIndex == cmbStatusList.Items.Count - 1)
            {
                //Проходим всех преподавателей и убираем у них
                //значение удаляемого звания
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Status == null))
                    {
                        if (mdlData.colLecturer[i].Status.Status.Equals(mdlData.colStatus[cmbStatusList.SelectedIndex].Status))
                        {
                            mdlData.colLecturer[i].Status = null;
                        }
                    }
                }

                //Удаляем степень из коллекции степеней
                mdlData.colStatus.RemoveAt(mdlData.colStatus.Count - 1);
                //Удаляем степень из списка степеней
                cmbStatusList.Items.RemoveAt(cmbStatusList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbStatusList.SelectedIndex = cmbStatusList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbStatusList.SelectedIndex;
                //Проходим всех преподавателей и убираем у них
                //значение удаляемого звания
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Status == null))
                    {
                        if (mdlData.colLecturer[i].Status.Status.Equals(mdlData.colStatus[DelElem].Status))
                        {
                            mdlData.colLecturer[i].Status = null;
                        }
                    }
                }

                mdlData.colStatus.RemoveAt(DelElem);
                cmbStatusList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colStatus.Count - 1; i++)
                {
                    mdlData.colStatus[i].Code = mdlData.colStatus[i].Code - 1;
                }
                cmbStatusList.SelectedIndex = DelElem;
                FillStatusList();
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Звание"
            clsStatus St = new clsStatus();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            St.Code = mdlData.colStatus.Count + 1;
            //Формируем задел названия для новой степени
            St.Status = "Новое звание";
            //Добавляем объект в коллекцию
            mdlData.colStatus.Add(St);
            //Заносим объект в список
            cmbStatusList.Items.Add(mdlData.colStatus[mdlData.colStatus.Count - 1].Code + ". " + mdlData.colStatus[mdlData.colStatus.Count - 1].Status);
            //Переходим к новому элементу списка
            cmbStatusList.SelectedIndex = cmbStatusList.Items.Count - 1;
        }

        private void cmbStatusList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtStatus.Text = mdlData.colStatus[cmbStatusList.SelectedIndex].Status;
        }
    }
}
