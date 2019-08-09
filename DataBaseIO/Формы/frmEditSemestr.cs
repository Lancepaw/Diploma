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
    public partial class frmEditSemestr : Form
    {
        public frmEditSemestr()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditSemestr_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы семестров
            if (mdlData.colSemestr.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillSemestrList();

            }
            //при неудачной загрузке элементов из таблицы семестров
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillSemestrList()
        {
            int NumFix = 0;
            NumFix = cmbSemestrList.SelectedIndex;
            //Очищаем список
            cmbSemestrList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestrList.Items.Add(mdlData.colSemestr[i].Code + ". " + mdlData.colSemestr[i].SemNum);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSemestrList.SelectedIndex = 0;
            }
            else
            {
                cmbSemestrList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colSemestr.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbSemestrList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtSemNum.Text = mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Семестр"
            clsSemestr Sem = new clsSemestr();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Sem.Code = mdlData.colSemestr.Count + 1;
            //Формируем задел названия для нового семестра
            Sem.SemNum = "i семестр";
            //Добавляем объект в коллекцию
            mdlData.colSemestr.Add(Sem);
            //Заносим объект в список
            cmbSemestrList.Items.Add(mdlData.colSemestr[mdlData.colSemestr.Count - 1].Code + ". " + mdlData.colSemestr[mdlData.colSemestr.Count - 1].SemNum);
            //Переходим к новому элементу списка
            cmbSemestrList.SelectedIndex = cmbSemestrList.Items.Count - 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum = txtSemNum.Text;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbSemestrList.SelectedIndex == cmbSemestrList.Items.Count - 1)
            {
                //Проходим всю нагрузку
                //и убираем значение удаляемого семестра
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Semestr == null))
                    {
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum))
                        {
                            mdlData.colDistribution[i].Semestr = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Semestr == null))
                    {
                        if (mdlData.colHouredDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum))
                        {
                            mdlData.colHouredDistribution[i].Semestr = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Semestr == null))
                    {
                        if (mdlData.colCombineDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum))
                        {
                            mdlData.colCombineDistribution[i].Semestr = null;
                        }
                    }
                }

                //Удаляем семестр из коллекции семестров
                mdlData.colSemestr.RemoveAt(mdlData.colSemestr.Count - 1);
                //Удаляем семестр из списка семестров
                cmbSemestrList.Items.RemoveAt(cmbSemestrList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbSemestrList.SelectedIndex = cmbSemestrList.Items.Count - 1;
            }

            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbSemestrList.SelectedIndex;
                //Проходим всю нагрузку и убираем
                //значение удаляемого семестра
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Semestr == null))
                    {
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[DelElem].SemNum))
                        {
                            mdlData.colDistribution[i].Semestr = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Semestr == null))
                    {
                        if (mdlData.colHouredDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[DelElem].SemNum))
                        {
                            mdlData.colHouredDistribution[i].Semestr = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Semestr == null))
                    {
                        if (mdlData.colCombineDistribution[i].Semestr.SemNum.Equals(mdlData.colSemestr[DelElem].SemNum))
                        {
                            mdlData.colCombineDistribution[i].Semestr = null;
                        }
                    }
                }

                mdlData.colSemestr.RemoveAt(DelElem);
                cmbSemestrList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colSemestr.Count - 1; i++)
                {
                    mdlData.colSemestr[i].Code = mdlData.colSemestr[i].Code - 1;
                }

                cmbSemestrList.SelectedIndex = DelElem;
                FillSemestrList();
            }
        }
    }
}
