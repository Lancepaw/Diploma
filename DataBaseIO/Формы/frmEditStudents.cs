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
    public partial class frmEditStudents : Form
    {
        private int MaxNum;

        public frmEditStudents()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditStudents_Load(object sender, EventArgs e)
        {
            FillKursList();
            FillLecturerList(cmbLecturerList);
            FillLecturerList(cmbLecturerFilt);
            FillSpecialityList();
            FillDepartmentList();

            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы дисциплин
            if (mdlData.colStudents.Count > 0)
            {
                btnSave.Enabled = true;
                btnCopy.Enabled = true;
                btnDel.Enabled = true;
                FillStudentsList(cmbStudentsList, mdlData.colStudents, false);

                cmbLecturerFilt.Enabled = mdlData.flgStudLecturerFilt;
                cmbLecturerFilt.SelectedIndex = mdlData.inxStudLecturer;
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnCopy.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        //
        private void FillDepartmentList()
        {
            //Очищаем список
            cmbDepartmentList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                cmbDepartmentList.Items.Add(mdlData.colDepart[i].Code + ". " + 
                    mdlData.colDepart[i].Short);
            }

            cmbDepartmentList.SelectedIndex = -1;
        }

        private void FillSpecialityList()
        {
            //Очищаем список
            cmbSpecialityList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                cmbSpecialityList.Items.Add(mdlData.colSpecialisation[i].Code + ". " + 
                    mdlData.colSpecialisation[i].ShortUpravlenie +
                    " [" + mdlData.colSpecialisation[i].ShortDop + "]" +
                    " (" + mdlData.colSpecialisation[i].ShortInstitute + ")");
            }

            cmbSpecialityList.SelectedIndex = -1;
        }

        private void FillKursList()
        {
            //Очищаем список
            cmbKursList.Items.Clear();

            //Заполняем комбо-список номерами курсов
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                cmbKursList.Items.Add(mdlData.colKursNum[i].Kurs);
            }

            cmbKursList.SelectedIndex = -1;
        }

        private void FillLecturerList(ComboBox cmb)
        {
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список преподавателями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }

            cmb.SelectedIndex = -1;
        }

        private void FillStudentsList(ComboBox cmb,
            IList<clsStudents> collection,
            bool flgLecturerFilt)
        {

            bool NeedFilt = false;
            bool AddToColl = false;

            int i;

            //Очистка комбо-списка студентов
            cmb.Items.Clear();

            //Если коллекция отфильтрованных студентов не пуста
            if (!(mdlData.FiltredStudents == null))
            {
                //то очищаем эту коллекцию
                mdlData.FiltredStudents.Clear();
            }

            //Если пришедшая в процедуру основная коллекция не пуста,
            //то начинаем работу по набору фильтрованных элементов
            if (!(collection == null))
            {
                //Сбрасываем сведения о максимальном коде элемента в коллекции
                MaxNum = 0;
                //Заполняем комбо-список студентами
                for (i = 0; i <= collection.Count - 1; i++)
                {
                    //По умолчанию считаем, что элемент нужно добавить в коллекцию
                    AddToColl = true;
                    
                    //Если максимальный код оказался меньше кода поступившего элемента,
                    //значит найденный ранее код не максимален - заменяем его
                    if (MaxNum < collection[i].Code)
                    {
                        MaxNum = collection[i].Code;
                    }

                    //Определяем, нужно ли фильтровать согласно флагам
                    if (flgLecturerFilt || false ||
                        false || false ||
                        false || false ||
                        false)
                    {
                        NeedFilt = true;
                    }

                    //Если фильтровать не нужно, то добавляем всё
                    if (!(NeedFilt))
                    {
                        cmbStudentsList.Items.Add(collection[i].Code + ". " + collection[i].FIO);
                    }
                    else
                    {
                        //Если требуется фильтровать по руководителю
                        if (flgLecturerFilt)
                        {
                            //Если преподаватель указан надо ещё принять решение
                            if (!(collection[i].Lect == null))
                            {
                                //Если в перечне выбранный индекс не пуст
                                if (cmbLecturerFilt.SelectedIndex > -1)
                                {
                                    //Если преподаватели совпали, то нас интересует такой элемент
                                    if (collection[i].Lect.FIO.Equals(mdlData.colLecturer[cmbLecturerFilt.SelectedIndex].FIO))
                                    {
                                        AddToColl = (true & AddToColl);
                                    }
                                    //Иначе не интересует
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если мы прошли все ступени и остались true,
                        //значит, мы соответствуем всем требованиям
                        //и должны добавить элемент в список
                        if (AddToColl)
                        {
                            cmb.Items.Add(collection[i].Code + ". " + collection[i].FIO);
                            mdlData.FiltredStudents.Add(collection[i]);
                        }
                    }
                }

                //
                if (cmb.Items.Count > 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = -1;
                }

                if (cmb.Equals(cmbStudentsList))
                {
                    txtRows.Text = (cmb.Items.Count).ToString();
                }
            }
        }

        private void CheckFiltParam(CheckBox chkCur, ComboBox cmbCur, int inx, ref bool flg)
        {
            if (chkCur.Checked)
            {
                cmbCur.Enabled = true;
                cmbCur.SelectedIndex = inx;
            }
            else
            {
                cmbCur.Enabled = false;
                cmbCur.SelectedIndex = -1;

                FillStudentsList(cmbStudentsList, mdlData.colStudents, chkLecturerFilt.Checked);
            }

            //Сохраняем глобальное состояние галочки
            flg = chkCur.Checked;
        }

        private void cmbStudentsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtTheme.Text = mdlData.colStudents[cmbStudentsList.SelectedIndex].Theme;
            txtFIO.Text = mdlData.colStudents[cmbStudentsList.SelectedIndex].FIO;

            chkInPlan.Checked = mdlData.colStudents[cmbStudentsList.SelectedIndex].flgPlan;

            chkHoured.Checked = mdlData.colStudents[cmbStudentsList.SelectedIndex].flgHoured;

            //Высвечивать номер курса
            if (!(mdlData.colStudents[cmbStudentsList.SelectedIndex].KursNum == null))
            {
                if (cmbKursList.Items.Count > 0)
                {
                    cmbKursList.SelectedIndex = mdlData.colKursNum.IndexOf(mdlData.colStudents[cmbStudentsList.SelectedIndex].KursNum);
                }
            }
            else
            {
                cmbKursList.SelectedIndex = -1;
            }

            //Высвечивать руководителя
            if (!(mdlData.colStudents[cmbStudentsList.SelectedIndex].Lect == null))
            {
                if (cmbLecturerList.Items.Count > 0)
                {
                    cmbLecturerList.SelectedIndex = mdlData.colLecturer.IndexOf(mdlData.colStudents[cmbStudentsList.SelectedIndex].Lect);
                }
            }
            else
            {
                cmbLecturerList.SelectedIndex = -1;
            }

            //Высвечивать кафедру
            if (!(mdlData.colStudents[cmbStudentsList.SelectedIndex].Depart == null))
            {
                if (cmbDepartmentList.Items.Count > 0)
                {
                    cmbDepartmentList.SelectedIndex = mdlData.colDepart.IndexOf(mdlData.colStudents[cmbStudentsList.SelectedIndex].Depart);
                }
            }
            else
            {
                cmbDepartmentList.SelectedIndex = -1;
            }

            //Высвечивать специальность
            if (!(mdlData.colStudents[cmbStudentsList.SelectedIndex].Speciality == null))
            {
                if (cmbSpecialityList.Items.Count > 0)
                {
                    cmbSpecialityList.SelectedIndex = mdlData.colSpecialisation.IndexOf(mdlData.colStudents[cmbStudentsList.SelectedIndex].Speciality);
                }
            }
            else
            {
                cmbSpecialityList.SelectedIndex = -1;
            }
        }

        //
        private void btnSave_Click(object sender, EventArgs e)
        {
            int NumFix = 0;
            
            mdlData.colStudents[cmbStudentsList.SelectedIndex].FIO = txtFIO.Text;
            mdlData.colStudents[cmbStudentsList.SelectedIndex].Theme = txtTheme.Text;

            mdlData.colStudents[cmbStudentsList.SelectedIndex].flgPlan = chkInPlan.Checked;

            mdlData.colStudents[cmbStudentsList.SelectedIndex].flgHoured = chkHoured.Checked;

            if (cmbDepartmentList.SelectedIndex >= 0)
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Depart = mdlData.colDepart[cmbDepartmentList.SelectedIndex];
            }
            else
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Depart = null;
            }

            if (cmbKursList.SelectedIndex >= 0)
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].KursNum = mdlData.colKursNum[cmbKursList.SelectedIndex];
            }
            else
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].KursNum = null;
            }

            if (cmbLecturerList.SelectedIndex >= 0)
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Lect = mdlData.colLecturer[cmbLecturerList.SelectedIndex];
            }
            else
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Lect = null;
            }

            if (cmbSpecialityList.SelectedIndex >= 0)
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Speciality = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex];
            }
            else
            {
                mdlData.colStudents[cmbStudentsList.SelectedIndex].Speciality = null;
            }

            NumFix = cmbStudentsList.SelectedIndex;

            FillStudentsList(cmbStudentsList, mdlData.colStudents, chkLecturerFilt.Checked);

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbStudentsList.SelectedIndex = 0;
            }
            else
            {
                cmbStudentsList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colStudents.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }

            mdlData.statString = "Последнее действие: Сохранение строки учебной нагрузки";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            btnCopy.Enabled = true;
            btnDel.Enabled = true;
            
            //Создаём новый объект класса "Студент"
            clsStudents Stud = new clsStudents();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Stud.Code = mdlData.colStudents.Count + 1;
            //Формируем задел названия для нового студента
            Stud.FIO = "Петров Петр Петрович";
            //По умолчанию студент включается в планирование нагрузки
            Stud.flgPlan = true;
            //Добавляем объект в коллекцию
            mdlData.colStudents.Add(Stud);
            //Заносим объект в список
            cmbStudentsList.Items.Add(mdlData.colStudents[mdlData.colStudents.Count - 1].Code + ". " + mdlData.colStudents[mdlData.colStudents.Count - 1].FIO);
            //Переходим к новому элементу списка
            cmbStudentsList.SelectedIndex = cmbStudentsList.Items.Count - 1;
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            int i;

            //Запрос на подтверждение пользователем удаления элемента
            if (MessageBox.Show(this, "Действительно удалить?", "Удаление",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                //Если удаляется последний элемент списка, то его можно
                //просто удалить
                if (cmbStudentsList.SelectedIndex == cmbStudentsList.Items.Count - 1)
                {
                    //Удаляем студента из коллекции студентов
                    mdlData.colStudents.RemoveAt(mdlData.colStudents.Count - 1);
                    //Удаляем студента из списка студентов
                    cmbStudentsList.Items.RemoveAt(cmbStudentsList.Items.Count - 1);
                    //Переводим выбранный индекс в списке на позицию вверх
                    cmbStudentsList.SelectedIndex = cmbStudentsList.Items.Count - 1;
                }
                //А если не последний, то нужно удалять аккуратно,
                //заменяя коды у всех элементов, оставшихся после
                else
                {
                    //Запоминаем индекс удаляемого элемента
                    DelElem = cmbStudentsList.SelectedIndex;

                    mdlData.colStudents.RemoveAt(DelElem);
                    cmbStudentsList.Items.RemoveAt(DelElem);

                    for (i = DelElem; i <= mdlData.colStudents.Count - 1; i++)
                    {
                        mdlData.colStudents[i].Code = mdlData.colStudents[i].Code - 1;
                    }

                    cmbStudentsList.SelectedIndex = DelElem;
                    FillStudentsList(cmbStudentsList, mdlData.colStudents, false);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Студент"
            clsStudents Stud = new clsStudents();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Stud.Code = mdlData.colStudents.Count + 1;
            //Формируем задел названия для нового студента
            Stud.FIO = txtFIO.Text + "+";
            //Учёт в планировании студента точно такой же, как и у выбранного
            Stud.flgPlan = chkInPlan.Checked;
            //Студент на той же кафедре, что и выбранный
            Stud.Depart = mdlData.colDepart[cmbDepartmentList.SelectedIndex];
            //Студент на том же курсе, что и выбранный
            Stud.KursNum = mdlData.colKursNum[cmbKursList.SelectedIndex];
            //Научный руководитель тот же, что и у выбранного
            Stud.Lect = mdlData.colLecturer[cmbLecturerList.SelectedIndex];
            //Специальность та же, что и у выбранного
            Stud.Speciality = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex];
            //Тему сбрасываем
            Stud.Theme = "";
            //Добавляем объект в коллекцию
            mdlData.colStudents.Add(Stud);
            //Заносим объект в список
            cmbStudentsList.Items.Add(mdlData.colStudents[mdlData.colStudents.Count - 1].Code + ". " + mdlData.colStudents[mdlData.colStudents.Count - 1].FIO);
            //Переходим к новому элементу списка
            cmbStudentsList.SelectedIndex = cmbStudentsList.Items.Count - 1;
        }

        private void cmbLecturerFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbLecturerFilt, ref mdlData.inxStudLecturer);
        }

        private void ChangeFiltParam(ComboBox cmbCur, ref int inx)
        {
            if (cmbCur.SelectedIndex >= 0)
            {
                //toolTip.SetToolTip(cmbCur,
                //    cmbCur.Items[cmbCur.SelectedIndex].ToString());
            }
            else
            {

            }

            inx = cmbCur.SelectedIndex;

            //Основной список заполняем только по отфильтрованной коллекции "Selected"
            FillStudentsList(cmbStudentsList, mdlData.colStudents, chkLecturerFilt.Checked);
        }

        private void chkLecturerFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkLecturerFilt, cmbLecturerFilt, mdlData.inxStudLecturer, ref mdlData.flgStudLecturerFilt);
        }
    }
}
