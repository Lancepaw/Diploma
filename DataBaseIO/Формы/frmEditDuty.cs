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
    public partial class frmEditDuty : Form
    {
        public frmEditDuty()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            //Закрыть форму редактирования должностей
            this.Close();
        }

        private void frmEditDuty_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы должностей
            if (mdlData.colDuty.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillDutyList();

            }
            //при неудачной загрузке элементов из таблицы должностей
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void cmbDutyList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtDutyFromList.Text = mdlData.colDuty[cmbDutyList.SelectedIndex].Duty;
            txtShort.Text = mdlData.colDuty[cmbDutyList.SelectedIndex].Short;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colDuty[cmbDutyList.SelectedIndex].Duty = txtDutyFromList.Text;
            mdlData.colDuty[cmbDutyList.SelectedIndex].Short = txtShort.Text;
            FillDutyList();
        }

        private void FillDutyList()
        {
            int NumFix = 0;

            NumFix = cmbDutyList.SelectedIndex;
            //Очищаем список
            cmbDutyList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDuty.Count - 1; i++)
            {
                cmbDutyList.Items.Add(mdlData.colDuty[i].Code + ". " + mdlData.colDuty[i].Duty);
            }

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbDutyList.SelectedIndex = 0;
            }
            else
            {
                cmbDutyList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colDuty.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Должностей"
            clsDuty D = new clsDuty();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            D.Code = mdlData.colDuty.Count + 1;
            //Формируем задел названия для новой должности
            D.Duty = "Новая должность";
            //Добавляем объект в коллекци
            mdlData.colDuty.Add(D);
            //Заносим объект в список
            cmbDutyList.Items.Add(mdlData.colDuty[mdlData.colDuty.Count-1].Code + ". " + mdlData.colDuty[mdlData.colDuty.Count-1].Duty);
            //Переходим к новому элементу списка
            cmbDutyList.SelectedIndex = cmbDutyList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbDutyList.SelectedIndex == cmbDutyList.Items.Count - 1)
            {
                //Проходим всех преподавателей и убираем у них
                //значение удаляемой должности
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Duty == null))
                    {
                        if (mdlData.colLecturer[i].Duty.Duty.Equals(mdlData.colDuty[cmbDutyList.SelectedIndex].Duty))
                        {
                            mdlData.colLecturer[i].Duty = null;
                        }
                    }
                    
                    if (!(mdlData.colLecturer[i].Duty1 == null))
                    {
                        if (mdlData.colLecturer[i].Duty1.Duty.Equals(mdlData.colDuty[cmbDutyList.SelectedIndex].Duty))
                        {
                            mdlData.colLecturer[i].Duty1 = null;
                        }
                    }
                }                  
                
                mdlData.colDuty.RemoveAt(mdlData.colDuty.Count - 1);
                cmbDutyList.Items.RemoveAt(cmbDutyList.Items.Count - 1);
                cmbDutyList.SelectedIndex = cmbDutyList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                DelElem = cmbDutyList.SelectedIndex;
                //Проходим всех преподавателей и убираем у них
                //значение удаляемой должности
                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if (!(mdlData.colLecturer[i].Duty == null))
                    {
                        if (mdlData.colLecturer[i].Duty.Duty.Equals(mdlData.colDuty[DelElem].Duty))
                        {
                            mdlData.colLecturer[i].Duty = null;
                        }
                    }
                    
                    if (!(mdlData.colLecturer[i].Duty1 == null))
                    {
                        if (mdlData.colLecturer[i].Duty1.Duty.Equals(mdlData.colDuty[DelElem].Duty))
                        {
                            mdlData.colLecturer[i].Duty1 = null;
                        }
                    }
                }

                mdlData.colDuty.RemoveAt(DelElem);
                cmbDutyList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colDuty.Count - 1; i++)
                {
                    mdlData.colDuty[i].Code = mdlData.colDuty[i].Code - 1;
                }
                
                cmbDutyList.SelectedIndex = DelElem;
                FillDutyList();
            }
        }
    }
}
