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
    public partial class frmEditSpeciality : Form
    {
        public frmEditSpeciality()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditSpeciality_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы специальностей
            if (mdlData.colSpecialisation.Count > 0)
            {
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                FillFacultyList();
                FillSpecialityList();
            }
            //при неудачной загрузке элементов из таблицы специальностей
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void FillSpecialityList()
        {
            int NumFix = 0;
            NumFix = cmbSpecialityList.SelectedIndex;
            //Очищаем список
            cmbSpecialityList.Items.Clear();

            //Заполняем комбо-список званиями
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                if (!mdlData.flgOldDB)
                {
                    cmbSpecialityList.Items.Add(mdlData.colSpecialisation[i].Code + ". " + mdlData.colSpecialisation[i].ShortUpravlenie +
                                                "-" + mdlData.colSpecialisation[i].ShortDop);
                }
                else
                {
                    cmbSpecialityList.Items.Add(mdlData.colSpecialisation[i].Code + ". " + mdlData.colSpecialisation[i].ShortUpravlenie);
                }
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSpecialityList.SelectedIndex = 0;
            }
            else
            {
                cmbSpecialityList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colSpecialisation.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }
        }

        private void FillFacultyList()
        {
            //Очищаем список
            cmbFaculty.Items.Clear();

            //Заполняем комбо-список званиями
            for (int i = 0; i <= mdlData.colFaculty.Count - 1; i++)
            {
                cmbFaculty.Items.Add(mdlData.colFaculty[i].Code + ". " + mdlData.colFaculty[i].Short);
            }
            
            cmbFaculty.SelectedIndex = -1;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Специализации"
            clsSpecialisation Sp = new clsSpecialisation();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Sp.Code = mdlData.colSpecialisation.Count + 1;
            //Формируем задел для новой (абстрактной) специализации
            Sp.Faculty = null;
            Sp.ShortDop = "----";
            Sp.ShortInstitute = "----";
            Sp.ShortUpravlenie = "----";
            Sp.Specialisation = "Новая специальность";
            Sp.Diff = "Б";
            //Добавляем объект в коллекцию
            mdlData.colSpecialisation.Add(Sp);
            //Заносим объект в список
            cmbSpecialityList.Items.Add(mdlData.colSpecialisation[mdlData.colSpecialisation.Count - 1].Code + ". " + mdlData.colSpecialisation[mdlData.colSpecialisation.Count - 1].ShortUpravlenie +
                                                "-" + mdlData.colSpecialisation[mdlData.colSpecialisation.Count - 1].ShortDop);
            //Переходим к новому элементу списка
            cmbSpecialityList.SelectedIndex = cmbSpecialityList.Items.Count - 1;
        }

        private void cmbSpecialityList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtSpecialityFull.Text = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Specialisation;
            txtSpecialityShort.Text = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortUpravlenie;
            txtSpecShort.Text = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortDop;
            txtInst.Text = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortInstitute;
            txtDiff.Text = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Diff;

            //Если факультет задан
            if (mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Faculty != null)
            {
                //Если количество элементов в списке больше нуля
                if (cmbFaculty.Items.Count > 0)
                {
                    cmbFaculty.SelectedIndex = mdlData.colFaculty.IndexOf(mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Faculty);
                }
            }
            else
            {
                cmbFaculty.SelectedIndex = -1;
            }
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Specialisation = txtSpecialityFull.Text;
            mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortUpravlenie = txtSpecialityShort.Text;
            mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortDop = txtSpecShort.Text;
            mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortInstitute = txtInst.Text;
            mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Diff = txtDiff.Text;

            if (cmbFaculty.SelectedIndex >= 0)
            {
                mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Faculty = mdlData.colFaculty[cmbFaculty.SelectedIndex];
            }
            else
            {
                mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].Faculty = null;
            }

            FillSpecialityList();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbSpecialityList.SelectedIndex == cmbSpecialityList.Items.Count - 1)
            {
                //Проходим всю нагрузку
                //и убираем удаляемую специальность отовсюду
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Speciality == null))
                    {
                        if (mdlData.colDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortDop))
                        {
                            mdlData.colDistribution[i].Speciality = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Speciality == null))
                    {
                        if (mdlData.colHouredDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortDop))
                        {
                            mdlData.colHouredDistribution[i].Speciality = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Speciality == null))
                    {
                        if (mdlData.colCombineDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex].ShortDop))
                        {
                            mdlData.colCombineDistribution[i].Speciality = null;
                        }
                    }
                }

                //Удаляем специальность из коллекции
                mdlData.colSpecialisation.RemoveAt(mdlData.colSpecialisation.Count - 1);
                //Удаляем специальность из списка
                cmbSpecialityList.Items.RemoveAt(cmbSpecialityList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbSpecialityList.SelectedIndex = cmbSpecialityList.Items.Count - 1;
            }

            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbSpecialityList.SelectedIndex;
                
                //Проходим всю нагрузку и убираем
                //отовсюду удаляемого преподавателя
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Speciality == null))
                    {
                        if (mdlData.colDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[DelElem].ShortDop))
                        {
                            mdlData.colDistribution[i].Speciality = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Speciality == null))
                    {
                        if (mdlData.colHouredDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[DelElem].ShortDop))
                        {
                            mdlData.colHouredDistribution[i].Speciality = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Speciality == null))
                    {
                        if (mdlData.colCombineDistribution[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[DelElem].ShortDop))
                        {
                            mdlData.colCombineDistribution[i].Speciality = null;
                        }
                    }
                }

                mdlData.colSpecialisation.RemoveAt(DelElem);
                cmbSpecialityList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colSpecialisation.Count - 1; i++)
                {
                    mdlData.colSpecialisation[i].Code = mdlData.colSpecialisation[i].Code - 1;
                }

                cmbSpecialityList.SelectedIndex = DelElem;
                FillSpecialityList();
            }
        }
    }
}
