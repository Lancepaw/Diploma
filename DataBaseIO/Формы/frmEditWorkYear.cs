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
    public partial class frmEditWorkYear : Form
    {
        public frmEditWorkYear()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear = txtWorkYear.Text;
            mdlData.statString = "Последнее действие: Изменение параметров учебного года";

            FillWorkYearList();
        }

        private void FillWorkYearList()
        {
            int NumFix = 0;

            NumFix = cmbWorkYearList.SelectedIndex;
            //Очищаем список
            cmbWorkYearList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colWorkYear.Count - 1; i++)
            {
                cmbWorkYearList.Items.Add(mdlData.colWorkYear[i].Code + ". " + mdlData.colWorkYear[i].WorkYear);
            }

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbWorkYearList.SelectedIndex = 0;
            }
            else
            {
                cmbWorkYearList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colWorkYear.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbWorkYearList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWorkYear.Text = mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear;
        }

        private void frmEditWorkYear_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из учебных годов
            if (mdlData.colWorkYear.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillWorkYearList();
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Учебный год"
            clsWorkYear WY = new clsWorkYear();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            WY.Code = mdlData.colWorkYear.Count + 1;
            //Формируем задел названия для нового учебного года
            WY.WorkYear = "----/----";
            //Добавляем объект в коллекцию
            mdlData.colWorkYear.Add(WY);
            //Заносим объект в список
            cmbWorkYearList.Items.Add(mdlData.colWorkYear[mdlData.colWorkYear.Count - 1].Code + ". " + mdlData.colWorkYear[mdlData.colWorkYear.Count - 1].WorkYear);
            //Переходим к новому элементу списка
            cmbWorkYearList.SelectedIndex = cmbWorkYearList.Items.Count - 1;

            mdlData.statString = "Последнее действие: Добавление нового учебного года";
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbWorkYearList.SelectedIndex == cmbWorkYearList.Items.Count - 1)
            {
                //Удаляем элемент из коллекции
                mdlData.colWorkYear.RemoveAt(mdlData.colWorkYear.Count - 1);
                //Удаляем элемент из списка
                cmbWorkYearList.Items.RemoveAt(cmbWorkYearList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbWorkYearList.SelectedIndex = cmbWorkYearList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbWorkYearList.SelectedIndex;

                mdlData.colWorkYear.RemoveAt(DelElem);
                cmbWorkYearList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colWorkYear.Count - 1; i++)
                {
                    mdlData.colWorkYear[i].Code = mdlData.colWorkYear[i].Code - 1;
                }
            }

            cmbWorkYearList.SelectedIndex = DelElem;
            mdlData.statString = "Последнее действие: Удаление учебного года";

            FillWorkYearList();
        }
    }
}