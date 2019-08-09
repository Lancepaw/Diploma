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
    public partial class frmEditSubjectTypes : Form
    {
        public frmEditSubjectTypes()
        {
            InitializeComponent();
        }

        private void frmEditSubjectTypes_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы видов учебной нагрузки
            if (mdlData.colSubject.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillSubjectTypesList();
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillSubjectTypesList()
        {
            int NumFix = 0;
            NumFix = cmbSubjectTypesList.SelectedIndex;
            //Очищаем список
            cmbSubjectTypesList.Items.Clear();

            //Заполняем комбо-список видами учебной нагрузки
            for (int i = 0; i <= mdlData.colSubjectType.Count - 1; i++)
            {
                cmbSubjectTypesList.Items.Add(mdlData.colSubjectType[i].Code + ". " + mdlData.colSubjectType[i].Type + " (" + mdlData.colSubjectType[i].Short + ")");
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSubjectTypesList.SelectedIndex = 0;
            }
            else
            {
                cmbSubjectTypesList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colSubjectType.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbSubjectTypesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtType.Text = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Type;
            txtShort.Text = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Short;
            txtPlan.Text = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].ShortPlan;
            txtDistrib.Text = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].LikeDistrib;
            txtForms.Text = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].ForForms;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Вид учебной нагрузки"
            clsSubjectType Type = new clsSubjectType();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Type.Code = mdlData.colSubjectType.Count + 1;
            //Формируем задел названия для нового вида учебной нагрузки
            Type.Type = "Новый вид";
            //Формируем задел короткого названия для вида учебной нагрузки
            Type.Short = "Нов.вид.";
            //Формируем задел короткого названия для вида учебной нагрузки
            //в индивидуальном плане
            Type.ShortPlan = "Нов.вид.";
            //Формируем задел короткого названия для вида учебной нагрузки
            //для таблицы распределения
            Type.LikeDistrib = "Нов_вид";
            //Формируем задел названия для вида учебной нагрузки
            //для печатных форм
            Type.ForForms = "Новый вид";
            //Добавляем объект в коллекцию
            mdlData.colSubjectType.Add(Type);
            //Заносим объект в список
            cmbSubjectTypesList.Items.Add(mdlData.colSubjectType[mdlData.colSubjectType.Count - 1].Code + ". " + mdlData.colSubjectType[mdlData.colSubjectType.Count - 1].Type + " (" + mdlData.colSubjectType[mdlData.colSubjectType.Count - 1].Short + ")");
            //Переходим к новому элементу списка
            cmbSubjectTypesList.SelectedIndex = cmbSubjectTypesList.Items.Count - 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Type = txtType.Text;
            mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Short = txtShort.Text;
            mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].ShortPlan = txtPlan.Text;
            mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].LikeDistrib = txtDistrib.Text;
            mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].ForForms = txtForms.Text;
            FillSubjectTypesList();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbSubjectTypesList.SelectedIndex == cmbSubjectTypesList.Items.Count - 1)
            {
                //Проходим все строки нагрузки и убираем в них
                //строки удаляемого вида нагрузки
                for (int i = 0; i <= mdlData.colDistributionDetailed.Count - 1; i++)
                {
                    if (!(mdlData.colDistributionDetailed[i].SubjType == null))
                    {
                        if (mdlData.colDistributionDetailed[i].SubjType.Short.Equals(mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Short))
                        {
                            mdlData.colDistributionDetailed.RemoveAt(i);
                            i--;
                        }
                    }
                }

                //Удаляем вид учебной нагрузки из коллекции дисциплин
                mdlData.colSubjectType.RemoveAt(mdlData.colSubjectType.Count - 1);
                //Удаляем вид учебной нагрузки из списка дисциплин
                cmbSubjectTypesList.Items.RemoveAt(cmbSubjectTypesList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbSubjectTypesList.SelectedIndex = cmbSubjectTypesList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbSubjectTypesList.SelectedIndex;
                //Проходим все строки нагрузки и убираем в них
                //значение удаляемой дисциплины
                for (int i = 0; i <= mdlData.colDistributionDetailed.Count - 1; i++)
                {
                    if (!(mdlData.colDistributionDetailed[i].SubjType == null))
                    {
                        if (mdlData.colDistributionDetailed[i].SubjType.Short.Equals(mdlData.colSubjectType[DelElem].Short))
                        {
                            mdlData.colDistributionDetailed.RemoveAt(i);
                            i--;
                        }
                    }
                }

                mdlData.colSubjectType.RemoveAt(DelElem);
                cmbSubjectTypesList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colSubjectType.Count - 1; i++)
                {
                    mdlData.colSubjectType[i].Code = mdlData.colSubjectType[i].Code - 1;
                }

                cmbSubjectTypesList.SelectedIndex = DelElem;
                FillSubjectTypesList();
            }
        }

        private void btnMoveUp_Click(object sender, EventArgs e)
        {
            clsSubjectType ST = new clsSubjectType();

            if (cmbSubjectTypesList.SelectedIndex > 1)
            {
                //У текущего элемента код уменьшить на единицу
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Code = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Code - 1;
                //У предыдущего элемента код увеличить на единицу
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex - 1].Code = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex - 1].Code + 1;

                ST = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex];
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex] = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex - 1];
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex - 1] = ST;
                
                //Смещаемся на предыдущий элемент
                cmbSubjectTypesList.SelectedIndex = cmbSubjectTypesList.SelectedIndex - 1;
                
                //Обновляем список
                FillSubjectTypesList();
            }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            clsSubjectType ST = new clsSubjectType();

            if (cmbSubjectTypesList.SelectedIndex < cmbSubjectTypesList.Items.Count - 1)
            {
                //У текущего элемента код увеличить на единицу
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Code = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex].Code + 1;
                //У следующего элемента код увеличить на единицу
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex + 1].Code = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex + 1].Code - 1;

                ST = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex];
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex] = mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex + 1];
                mdlData.colSubjectType[cmbSubjectTypesList.SelectedIndex + 1] = ST;

                //Смещаемся на следующий элемент
                cmbSubjectTypesList.SelectedIndex = cmbSubjectTypesList.SelectedIndex + 1;

                //Обновляем список
                FillSubjectTypesList();
            }
        }
    }
}
