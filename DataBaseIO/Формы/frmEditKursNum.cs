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
    public partial class frmEditKursNum : Form
    {
        public frmEditKursNum()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbKursNumList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtKursNum.Text = mdlData.colKursNum[cmbKursNumList.SelectedIndex].Kurs.ToString();
        }

        private void frmEditKursNum_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы номеров курсов
            if (mdlData.colKursNum.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillKursNumList();
            }
            //при неудачной загрузке элементов из таблицы номеров курсов
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillKursNumList()
        {
            int NumFix = 0;
            NumFix = cmbKursNumList.SelectedIndex;
            //Очищаем список
            cmbKursNumList.Items.Clear();

            //Заполняем комбо-список номерами курсов
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                cmbKursNumList.Items.Add(mdlData.colKursNum[i].Kurs);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbKursNumList.SelectedIndex = 0;
            }
            else
            {
                cmbKursNumList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colKursNum.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colKursNum[cmbKursNumList.SelectedIndex].Kurs = Convert.ToInt32(txtKursNum.Text);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Номер курса"
            clsKursNum KNum = new clsKursNum();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            KNum.Code = mdlData.colKursNum.Count + 1;
            //Формируем задел названия для нового номера курса
            KNum.Kurs = 0;
            //Добавляем объект в коллекцию
            mdlData.colKursNum.Add(KNum);
            //Заносим объект в список
            cmbKursNumList.Items.Add(mdlData.colKursNum[mdlData.colKursNum.Count - 1].Kurs);
            //Переходим к новому элементу списка
            cmbKursNumList.SelectedIndex = cmbKursNumList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbKursNumList.SelectedIndex == cmbKursNumList.Items.Count - 1)
            {
                //Проходим всю нагрузку
                //и убираем значение удаляемого курса
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].KursNum == null))
                    {
                        if (mdlData.colDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[cmbKursNumList.SelectedIndex].Kurs))
                        {
                            mdlData.colDistribution[i].KursNum = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].KursNum == null))
                    {
                        if (mdlData.colHouredDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[cmbKursNumList.SelectedIndex].Kurs))
                        {
                            mdlData.colHouredDistribution[i].KursNum = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].KursNum == null))
                    {
                        if (mdlData.colCombineDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[cmbKursNumList.SelectedIndex].Kurs))
                        {
                            mdlData.colCombineDistribution[i].KursNum = null;
                        }
                    }
                }

                //Удаляем факультет из коллекции номеров курсов
                mdlData.colKursNum.RemoveAt(mdlData.colKursNum.Count - 1);
                //Удаляем номер курса из списка номеров курсов
                cmbKursNumList.Items.RemoveAt(cmbKursNumList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbKursNumList.SelectedIndex = cmbKursNumList.Items.Count - 1;
            }
            
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbKursNumList.SelectedIndex;
                //Проходим всю нагрузку и убираем
                //значение удаляемого номера курса
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].KursNum == null))
                    {
                        if (mdlData.colDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[DelElem].Kurs))
                        {
                            mdlData.colDistribution[i].KursNum = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].KursNum == null))
                    {
                        if (mdlData.colHouredDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[DelElem].Kurs))
                        {
                            mdlData.colHouredDistribution[i].KursNum = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].KursNum == null))
                    {
                        if (mdlData.colCombineDistribution[i].KursNum.Kurs.Equals(mdlData.colKursNum[DelElem].Kurs))
                        {
                            mdlData.colCombineDistribution[i].KursNum = null;
                        }
                    }
                }

                mdlData.colKursNum.RemoveAt(DelElem);
                cmbKursNumList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colKursNum.Count - 1; i++)
                {
                    mdlData.colKursNum[i].Code = mdlData.colKursNum[i].Code - 1;
                }

                cmbKursNumList.SelectedIndex = DelElem;
                FillKursNumList();
            }
        }
    }
}
