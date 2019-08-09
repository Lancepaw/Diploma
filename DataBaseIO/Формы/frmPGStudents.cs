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
    public partial class frmPGStudents : Form
    {
        int MaxNum;
        
        public frmPGStudents()
        {
            InitializeComponent();
        }

        private void frmPGStudents_Load(object sender, EventArgs e)
        {
            FillKursList();
            FillLecturerList(cmbLecturerList);
            FillLecturerList(cmbLecturerFilt);
            FillDepartmentList();

            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы дисциплин
            if (mdlData.colStudents.Count > 0)
            {
                btnSave.Enabled = true;
                btnCopy.Enabled = true;
                btnDel.Enabled = true;
                FillStudentsList(cmbPGStudentsList, mdlData.colPGStudents, false);

                cmbLecturerFilt.Enabled = mdlData.flgSubjectFilt;
                cmbLecturerFilt.SelectedIndex = mdlData.inxSubject;
            }
            //при неудачной загрузке элементов из таблицы званий
            else
            {
                btnCopy.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FillDepartmentList()
        {
            //Очищаем список
            cmbDepartmentList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                cmbDepartmentList.Items.Add(mdlData.colDepart[i].Code + ". " + mdlData.colDepart[i].Short);
            }

            cmbDepartmentList.SelectedIndex = -1;
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
            IList<clsPGStudents> collection,
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
                        cmbPGStudentsList.Items.Add(collection[i].Code + ". " + collection[i].FIO);
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
                            mdlData.FiltredPGStudents.Add(collection[i]);
                        }
                    }
                }

                if (cmb.Items.Count > 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = -1;
                }

                if (cmb.Equals(cmbPGStudentsList))
                {
                    txtRows.Text = (cmb.Items.Count).ToString();
                }
            }
        }
    }
}
