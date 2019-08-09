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
    public partial class frmEditSubject : Form
    {
        public frmEditSubject()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditSubject_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы дисциплин
            if (mdlData.colSubject.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillSubjectList();
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillSubjectList()
        {
            int NumFix = 0;
            NumFix = cmbSubjectList.SelectedIndex;
            //Очищаем список
            cmbSubjectList.Items.Clear();

            //Заполняем комбо-список званиями
            for (int i = 0; i <= mdlData.colSubject.Count - 1; i++)
            {
                cmbSubjectList.Items.Add(mdlData.colSubject[i].Code + ". " + mdlData.colSubject[i].Subject);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSubjectList.SelectedIndex = 0;
            }
            else
            {
                cmbSubjectList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colSubject.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void cmbSubjectList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtSubjectName.Text = mdlData.colSubject[cmbSubjectList.SelectedIndex].Subject;
            txtSubjectShortName.Text = mdlData.colSubject[cmbSubjectList.SelectedIndex].SubjectShort;
            txtPreferences.Text = mdlData.colSubject[cmbSubjectList.SelectedIndex].Preferences;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Дисциплина"
            clsSubject Sub = new clsSubject();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Sub.Code = mdlData.colSubject.Count + 1;
            //Формируем задел названия для новой дисциплины
            Sub.Subject = "Новая дисциплина";
            //Формируем задел короткого названия для новой дисциплины
            Sub.SubjectShort = "Нов.дис.";
            //Формируем задел пожелания
            Sub.Preferences = "Аудитории ";
            //Добавляем объект в коллекцию
            mdlData.colSubject.Add(Sub);
            //Заносим объект в список
            cmbSubjectList.Items.Add(mdlData.colSubject[mdlData.colSubject.Count - 1].Code + ". " + mdlData.colSubject[mdlData.colSubject.Count - 1].Subject);
            //Переходим к новому элементу списка
            cmbSubjectList.SelectedIndex = cmbSubjectList.Items.Count - 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colSubject[cmbSubjectList.SelectedIndex].Subject = txtSubjectName.Text;
            mdlData.colSubject[cmbSubjectList.SelectedIndex].SubjectShort = txtSubjectShortName.Text;
            mdlData.colSubject[cmbSubjectList.SelectedIndex].Preferences = txtPreferences.Text;
            FillSubjectList();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbSubjectList.SelectedIndex == cmbSubjectList.Items.Count - 1)
            {
                //Проходим все строки нагрузки и убираем в них
                //значение удаляемой дисциплины
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (mdlData.colDistribution[i].Subject.Subject.Equals(mdlData.colSubject[cmbSubjectList.SelectedIndex].Subject))
                        {
                            mdlData.colDistribution[i].Subject = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Subject == null))
                    {
                        if (mdlData.colHouredDistribution[i].Subject.Subject.Equals(mdlData.colSubject[cmbSubjectList.SelectedIndex].Subject))
                        {
                            mdlData.colHouredDistribution[i].Subject = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Subject == null))
                    {
                        if (mdlData.colCombineDistribution[i].Subject.Subject.Equals(mdlData.colSubject[cmbSubjectList.SelectedIndex].Subject))
                        {
                            mdlData.colCombineDistribution[i].Subject = null;
                        }
                    }
                }

                //Удаляем дисциплину из коллекции дисциплин
                mdlData.colSubject.RemoveAt(mdlData.colSubject.Count - 1);
                //Удаляем дисциплину из списка дисциплин
                cmbSubjectList.Items.RemoveAt(cmbSubjectList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbSubjectList.SelectedIndex = cmbSubjectList.Items.Count - 1;
            }
            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbSubjectList.SelectedIndex;
                //Проходим все строки нагрузки и убираем в них
                //значение удаляемой дисциплины
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (mdlData.colDistribution[i].Subject.Subject.Equals(mdlData.colSubject[DelElem].Subject))
                        {
                            mdlData.colDistribution[i].Subject = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Subject == null))
                    {
                        if (mdlData.colHouredDistribution[i].Subject.Subject.Equals(mdlData.colSubject[DelElem].Subject))
                        {
                            mdlData.colHouredDistribution[i].Subject = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Subject == null))
                    {
                        if (mdlData.colCombineDistribution[i].Subject.Subject.Equals(mdlData.colSubject[DelElem].Subject))
                        {
                            mdlData.colCombineDistribution[i].Subject = null;
                        }
                    }
                }

                mdlData.colSubject.RemoveAt(DelElem);
                cmbSubjectList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colSubject.Count - 1; i++)
                {
                    mdlData.colSubject[i].Code = mdlData.colSubject[i].Code - 1;
                }

                cmbSubjectList.SelectedIndex = DelElem;
                FillSubjectList();
            }
        }
    }
}
