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
    public partial class frmEditCombination : Form
    {
        public frmEditCombination()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colCombination[cmbCombinationList.SelectedIndex].CombType = txtCombination.Text;
            FillCombinationList();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Совместительство"
            clsCombination Com = new clsCombination();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Com.Code = mdlData.colCombination.Count + 1;
            //Формируем задел названия для нового совместительства
            Com.CombType = "Новое совместительство";
            //Добавляем объект в коллекцию
            mdlData.colCombination.Add(Com);
            //Заносим объект в список
            cmbCombinationList.Items.Add(mdlData.colCombination[mdlData.colCombination.Count - 1].Code + ". " + mdlData.colCombination[mdlData.colCombination.Count - 1].CombType);
            //Переходим к новому элементу списка
            cmbCombinationList.SelectedIndex = cmbCombinationList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;

            //Запрос на подтверждение пользователем удаления элемента
            if (MessageBox.Show(this, "Действительно удалить?", "Удаление",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                //Если удаляется последний элемент списка, то его можно
                //просто удалить
                if (cmbCombinationList.SelectedIndex == cmbCombinationList.Items.Count - 1)
                {
                    //Проходим всех преподавателей и убираем у них
                    //значение удаляемого типа совместительства
                    for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                    {
                        if (!(mdlData.colLecturer[i].Combination == null))
                        {
                            if (mdlData.colLecturer[i].Combination.CombType.Equals(mdlData.colCombination[cmbCombinationList.SelectedIndex].CombType))
                            {
                                mdlData.colLecturer[i].Combination = null;
                            }
                        }
                    }
                    //Удаляем тип совместительства из коллекции совместительств
                    mdlData.colCombination.RemoveAt(mdlData.colCombination.Count - 1);
                    //Удаляем тип совместительства из списка совместительств
                    cmbCombinationList.Items.RemoveAt(cmbCombinationList.Items.Count - 1);
                    //Переводим выбранный индекс в списке на позицию вверх
                    cmbCombinationList.SelectedIndex = cmbCombinationList.Items.Count - 1;
                }
                //А если не последний, то нужно удалять аккуратно,
                //заменяя коды у всех элементов, оставшихся после
                else
                {
                    //Запоминаем индекс удаляемого элемента
                    DelElem = cmbCombinationList.SelectedIndex;
                    //Проходим всех преподавателей и убираем у них
                    //значение удаляемого типа совместительства
                    for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                    {
                        if (!(mdlData.colLecturer[i].Combination == null))
                        {
                            if (mdlData.colLecturer[i].Combination.CombType.Equals(mdlData.colCombination[DelElem].CombType))
                            {
                                mdlData.colLecturer[i].Combination = null;
                            }
                        }
                    }

                    mdlData.colCombination.RemoveAt(DelElem);
                    cmbCombinationList.Items.RemoveAt(DelElem);

                    for (int i = DelElem; i <= mdlData.colCombination.Count - 1; i++)
                    {
                        mdlData.colCombination[i].Code = mdlData.colCombination[i].Code - 1;
                    }

                    cmbCombinationList.SelectedIndex = DelElem;
                    FillCombinationList();
                }
            }
        }

        private void cmbCombinationList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtCombination.Text = mdlData.colCombination[cmbCombinationList.SelectedIndex].CombType;
        }

        private void frmEditCombination_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы совместительств
            if (mdlData.colDegree.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillCombinationList();
            }
            //при неудачной загрузке элементов из таблицы совместительств
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillCombinationList()
        {
            int NumFix = 0;
            NumFix = cmbCombinationList.SelectedIndex;
            //Очищаем список
            cmbCombinationList.Items.Clear();

            //Заполняем комбо-список степенями
            for (int i = 0; i <= mdlData.colDegree.Count - 1; i++)
            {
                cmbCombinationList.Items.Add(mdlData.colCombination[i].Code + ". " + mdlData.colCombination[i].CombType);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbCombinationList.SelectedIndex = 0;
            }
            else
            {
                cmbCombinationList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colCombination.Count == 1)
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
