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
    public partial class frmEditLecturer : Form
    {
        private int curNum;        
        
        public frmEditLecturer()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditLecturer_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы преподавателей
            frCount.Enabled = false;
            frDoneList.Enabled = false;
            lblAge.Enabled = false;
            txtAge.Enabled = false;
            lblSeniority.Enabled = false;
            txtSeniority.Enabled = false;
            txtSumRate.Enabled = false;
            txtSumUnload.Enabled = false;
            txtUpLoad.Enabled = false;
            txtRateWOUnload.Enabled = false;
            txtSumOldRate.Enabled = false;
            
            if (mdlData.colLecturer.Count > 0)
            {
                //Делаем доступными кнопки редактирования элементов
                btnAdd.Enabled = true;
                btnSave.Enabled = true;
                btnDel.Enabled = true;
                //Делаем доступными комбо-боксы
                cmbDutyAddList.Enabled = true;
                cmbDutyList.Enabled = true;
                //Заполняем комбо-боксы
                FillDutyList(cmbDutyList);
                FillDutyList(cmbDutyAddList);
                FillDepartmentList();
                FillStatusList();
                FillCombinationList();
                FillDegreeList();

                optCombine.Checked = false;
                optHoured.Checked = false;
                optMain.Checked = true;
            }
            //при неудачной загрузке элементов из таблицы должностей
            else
            {
                btnAdd.Enabled = false;
                btnSave.Enabled = false;
                btnDel.Enabled = false;
                cmbDutyAddList.Enabled = false;
                cmbDutyList.Enabled = false;
            }
        }

        private void FillLecturerList()
        {
            int NumFix = 0;
            double SumRate = 0;
            double RealRate = 0;
            double SumRateOld = 0;
            int sumUnLoad = 0;
            double sumNotUnLoadRate = 0;

            NumFix = cmbLecturerList.SelectedIndex;
            //Очищаем список
            cmbLecturerList.Items.Clear();

            //Заполняем комбо-список должностями
            //и попутно считаем суммарную ставку
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturerList.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);

                if (mdlData.colLecturer[i].ChangeRate)
                {
                    RealRate = (mdlData.colLecturer[i].Rate1 + mdlData.colLecturer[i].Rate2) / 2;
                }
                else
                {
                    RealRate = mdlData.colLecturer[i].Rate;
                }

                SumRate += mdlData.colLecturer[i].Rate;
                SumRateOld += mdlData.colLecturer[i].OldRate;

                //Считаем догрузку для преподавателей без разгрузки
                if (mdlData.colLecturer[i].UnLoad == 0)
                {
                    sumNotUnLoadRate += mdlData.colLecturer[i].Rate;
                }
                else
                {
                    sumUnLoad += mdlData.colLecturer[i].UnLoad;
                }
            }

            mdlData.LoadInc = Convert.ToInt32(-sumUnLoad / sumNotUnLoadRate);

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbLecturerList.SelectedIndex = 0;
            }
            else
            {
                cmbLecturerList.SelectedIndex = NumFix;
            }

            //Запрет на удаление единственного элемента из коллекции
            if (mdlData.colLecturer.Count == 1)
            {
                btnDel.Enabled = false;
            }
            else
            {
                btnDel.Enabled = true;
            }

            txtSumRate.Text = SumRate.ToString();
            txtSumOldRate.Text = SumRateOld.ToString();
            txtSumUnload.Text = (-sumUnLoad).ToString();
            txtRateWOUnload.Text = sumNotUnLoadRate.ToString();
            txtUpLoad.Text = mdlData.LoadInc.ToString();
        }

        private void FillDepartmentList()
        {
            int NumFix = 0;
            NumFix = cmbDepartmentList.SelectedIndex;
            //Очищаем список
            cmbDepartmentList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                cmbDepartmentList.Items.Add(mdlData.colDepart[i].Code + ". " + mdlData.colDepart[i].Short);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbDepartmentList.SelectedIndex = 0;
            }
            else
            {
                cmbDepartmentList.SelectedIndex = NumFix;
            }
        }

        private void FillCombinationList()
        {
            int NumFix = 0;
            NumFix = cmbCombinationList.SelectedIndex;
            //Очищаем список
            cmbCombinationList.Items.Clear();

            //Заполняем комбо-список совместительствами
            for (int i = 0; i <= mdlData.colCombination.Count - 1; i++)
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
        }

        private void FillStatusList()
        {
            int NumFix = 0;
            NumFix = cmbStatusList.SelectedIndex;
            //Очищаем список
            cmbStatusList.Items.Clear();

            //Заполняем комбо-список совместительствами
            for (int i = 0; i <= mdlData.colStatus.Count - 1; i++)
            {
                cmbStatusList.Items.Add(mdlData.colStatus[i].Code + ". " + mdlData.colStatus[i].Status);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbStatusList.SelectedIndex = 0;
            }
            else
            {
                cmbStatusList.SelectedIndex = NumFix;
            }
        }

        private void FillDegreeList()
        {
            int NumFix = 0;
            NumFix = cmbDegreeList.SelectedIndex;
            //Очищаем список
            cmbDegreeList.Items.Clear();

            //Заполняем комбо-список совместительствами
            for (int i = 0; i <= mdlData.colDegree.Count - 1; i++)
            {
                cmbDegreeList.Items.Add(mdlData.colDegree[i].Code + ". " + mdlData.colDegree[i].Short);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbDegreeList.SelectedIndex = 0;
            }
            else
            {
                cmbDegreeList.SelectedIndex = NumFix;
            }
        }

        private void FillDutyList(ComboBox cmb)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colDuty.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colDuty[i].Code + ". " + mdlData.colDuty[i].Duty);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmb.SelectedIndex = 0;
            }
            else
            {
                cmb.SelectedIndex = NumFix;
            }
        }

        /// <summary>
        /// Действия, происходящие при смене элемента в списке преподавателей
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbLecturerList_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Значение средней нагрузки на кафедру
            int AvgLoad = mdlData.AverageLoad;
            //Разгрузка дополнительная фактическая
            double UnLoadDopFact = 0;
            //Фактические часы нагрузки преподавателя
            int FactHours = 0;
            //Плановые часы нагрузки преподавателя
            double PlanHours = 0;
            //Реальная ставка при смене ставок в семестрах
            double RealRate = 0;

            //Возраст
            int Age = 0;
            //Стаж работы
            int Seniority = 0;

            IList<clsDistribution> coll;

            string[] FIO;
            
            //Заполняем строку должности
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty == null))
            {
                txtDuty.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty.Duty;
            }
            else
            {
                txtDuty.Text = "";
            }
            //Заполняем строку дополнительной должности
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1 == null))
            {
                txtDopDuty.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1.Duty;
            }
            else
            {
                txtDopDuty.Text = "";
            }
            //Заполняем строку кафедры
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Depart == null))
            {
                txtDepart.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Depart.Short;
            }
            else
            {
                txtDepart.Text = "";
            }
            //Заполняем строку учёной степени
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree == null))
            {
                txtDegree.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree.Short;
            }
            else
            {
                txtDegree.Text = "";
            }
            //Заполняем строку учёного звания
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status == null))
            {
                txtStatus.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status.Status;
            }
            else
            {
                txtStatus.Text = "";
            }
            //Заполняем строку совместительства
            if (!(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination == null))
            {
                txtCombination.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination.CombType;
            }
            else
            {
                txtCombination.Text = "";
            }

            //Установка соответствующего индекса должности
            cmbDutyList.SelectedIndex = mdlData.colDuty.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty);
            
            //Установка соответствующего индекса дополнительной должности
            cmbDutyAddList.SelectedIndex = mdlData.colDuty.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1);
            
            //Установка соответствующего индекса кафедры
            cmbDepartmentList.SelectedIndex = mdlData.colDepart.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Depart);
            
            //Установка соответствующего индекса совместительства
            cmbCombinationList.SelectedIndex = mdlData.colCombination.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination);
            
            //Установка соответствующего индекса степени
            cmbDegreeList.SelectedIndex = mdlData.colDegree.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree);
            
            //Установка соответствующего индекса звания
            cmbStatusList.SelectedIndex = mdlData.colStatus.IndexOf(mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status);
            
            //Формируем читаемую дату рождения
            txtDOB.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB.ToShortDateString();
            
            //Считаем возраст преподавателя
            Age = DateTime.Now.Year - mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB.Year;
            
            //Если текущий месяц не достиг месяца рождения, то необходимо
            //убавить один год в возрасте
            if (DateTime.Now.Month < mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB.Month)
            {
                Age -= 1;
            }
            //в противном случае, 
            else
            {
                //если месяц равен месяцу рождения
                if (DateTime.Now.Month == mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB.Month)
                {
                    //Если текущий день не достиг дня рождения, то необходимо
                    //убавить один год в возрасте
                    if (DateTime.Now.Day < mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB.Day)
                    {
                        Age -= 1;
                    }
                }
            }
            
            //Выводим возраст преподавателя
            txtAge.Text = Age.ToString();

            //Формируем читаемую дату начала работы
            txtWorkSince.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority.ToShortDateString();
            
            //Считаем стаж преподавателя
            Seniority = DateTime.Now.Year - mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority.Year;
            
            //Если текущий месяц не достиг месяца вступления на работу, то необходимо
            //убавить один год стажа
            if (DateTime.Now.Month < mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority.Month)
            {
                Seniority -= 1;
            }
            //в противном случае, 
            else
            {
                //если месяц равен месяцу начала работы
                if (DateTime.Now.Month == mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority.Month)
                {
                    //Если текущий день не достиг дня начала работы, то необходимо
                    //убавить один год стажа
                    if (DateTime.Now.Day < mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority.Day)
                    {
                        Seniority -= 1;
                    }
                }
            }
            
            //Выводим стаж преподавателя
            txtSeniority.Text = Seniority.ToString();
            
            //Выводим максимальную нагрузку преподавателя
            txtMaxLoad.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].MaxLoad.ToString();
            
            //Выводим разгрузку преподавателя
            txtUnLoad.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].UnLoad.ToString();

            //Разбираем строку для вывода отдельно
            //Фамилии, имени и отчества преподавателя
            FIO = mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO.Split(new char[] {' '});

            if (FIO.GetLength(0) == 3)
            {
                txtSurname.Text = FIO[0];
                txtName.Text = FIO[1];
                txtPatronymic.Text = FIO[2];
            }
            else if (FIO.GetLength(0) == 2)
            {
                txtSurname.Text = FIO[0];
                txtName.Text = FIO[1];
                txtPatronymic.Text = "";
            }
            else if (FIO.GetLength(0) == 1)
            {
                txtSurname.Text = FIO[0];
                txtName.Text = "";
                txtPatronymic.Text = "";
            }
            else
            {
                txtSurname.Text = "";
                txtName.Text = "";
                txtPatronymic.Text = "";
            }

            if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].ChangeRate)
            {
                RealRate = (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate1 +
                    mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate2) / 2;
            }
            else
            {
                RealRate = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate;
            }

            //Выводим реальную ставку преподавателя
            txtRealRate.Text = RealRate.ToString("0.00");

            //Выводим ставку преподавателя
            txtRate.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate.ToString("0.00");

            //Выводим нагрузку преподавателя
            //Здесь нужна реальная ставка
            if (optCombine.Checked || optMain.Checked)
            {
                txtLoad.Text = (RealRate * AvgLoad).ToString("0.00");
            }
            else
            {
                txtLoad.Text = "0";
            }
            
            //Считаем дополнительную нагрузку преподавателя с
            //учётом разгрузки других (здесь нужна ставка по отделу кадров)
            if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].UnLoad == 0 & !chkNonOverLoad.Checked)
            {
                UnLoadDopFact = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate * mdlData.LoadInc;
            }
            else
            {
                UnLoadDopFact = 0;
            }
            //Пересчитываем нагрузку преподавателя с учётом дополнительной и реальной
            if (optCombine.Checked || optMain.Checked)
            {
                PlanHours = (RealRate * AvgLoad) +
                             mdlData.colLecturer[cmbLecturerList.SelectedIndex].UnLoad + UnLoadDopFact;
            }
            else
            {
                PlanHours = 0;
            }
            
            //Пересчитанная нагрузка является плановой
            txtPlanLoad.Text = PlanHours.ToString("0.00");

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                else
                {
                    if (optCombine.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                    }
                    else
                    {
                        coll = mdlData.colDistribution;
                    }
                }
            }
            
            //Считаем фактическую нагрузку преподавателя
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если строка участвует в расчёте нагрузки
                if (!coll[i].flgExclude)
                {
                    //Если стандартная строка нагрузки
                    if (!coll[i].flgDistrib)
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                            {
                                FactHours += mdlData.toSumDistributionComponents(coll[i]);
                            }
                        }
                    }
                    //Если равномерно распределяемая строка нагрузки
                    else
                    {
                        //Проходим всех студентов студентов
                        for (int j = 0; j <= mdlData.colStudents.Count - 1; j++)
                        {
                            //Если студент записан в плановую нагрузку
                            if (mdlData.colStudents[j].flgPlan)
                            {
                                //Если рассматриваемый преподаватель - руководитель студента
                                //И если студент на том же курсе, где и дисциплина
                                //И специальность студента должна соответствовать специальности нагрузки
                                if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                    & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                    & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                {
                                    FactHours += coll[i].Weight;
                                }
                            }
                        }

                        //Может оказаться так, что часть равномерно распределяемой нагрузки
                        //переведено в почасовую и это необходимо учесть при расчёте
                        if (optCombine.Checked)
                        {
                            mdlData.toDetectUniformInHoured(ref FactHours, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                        }
                    }
                }
            }            
            //Выводим фактическую нагрузку преподавателя
            txtFactLoad.Text = FactHours.ToString("0.00");
            //Выводим перегрузку (+) или недогрузку (-) преподавателя
            txtOverLoad.Text = (FactHours - PlanHours).ToString("0.00");

            //Отмечаем членство в совете
            chkSoviet.Checked = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Soviet;
            //Выводим пожелания для диспетчерской
            txtText.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Text;
            //Выводим пожелания для диспетчерской
            txtPreferences.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Preferences;
            //Выводим старую ставку
            txtOldRate.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].OldRate.ToString("0.00");
            //Выводим ставку первого семестра
            txtRate1.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate1.ToString("0.00");
            //Выводим ставку второго семестра
            txtRate2.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate2.ToString("0.00");
            //Отмечаем необходимость смены ставок
            chkChangeRate.Checked = mdlData.colLecturer[cmbLecturerList.SelectedIndex].ChangeRate;

        }

        private void optMain_CheckedChanged(object sender, EventArgs e)
        {
            if (optMain.Checked)
            {
                if (cmbLecturerList.SelectedIndex >= 0)
                {
                    curNum = cmbLecturerList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }

                FillLecturerList();

                if (curNum <= cmbLecturerList.Items.Count - 1)
                {
                    cmbLecturerList.SelectedIndex = curNum;
                }

                GlobalLoadCount();
            }
        }

        private void optHoured_CheckedChanged(object sender, EventArgs e)
        {
            if (optHoured.Checked)
            {
                if (cmbLecturerList.SelectedIndex >= 0)
                {
                    curNum = cmbLecturerList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }
                FillLecturerList();
                if (curNum <= cmbLecturerList.Items.Count - 1)
                {
                    cmbLecturerList.SelectedIndex = curNum;
                }
                GlobalLoadCount();
            }
        }

        private void optCombine_CheckedChanged(object sender, EventArgs e)
        {
            if (optCombine.Checked)
            {
                if (cmbLecturerList.SelectedIndex >= 0)
                {
                    curNum = cmbLecturerList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }
                FillLecturerList();
                if (curNum <= cmbLecturerList.Items.Count - 1)
                {
                    cmbLecturerList.SelectedIndex = curNum;
                }
                GlobalLoadCount();
            }
        }

        private void GlobalLoadCount()
        {
            //Сумма перегрузки
            double PositiveSum = 0;
            //Сумма недогрузки
            double NegativeSum = 0;
            //Разница
            double Differ = 0;
            //Значение средней нагрузки на кафедру
            int AvgLoad = 750;
            //Значение средней нагрузки, распеределяемой между преподавателями
            //без разгрузки, отрабатывающих разгрузку других
            int UnLoadDopValue = 80;
            //Разгрузка дополнительная фактическая
            double UnLoadDopFact = 0;
            //Фактические часы нагрузки преподавателя
            int FactHours = 0;
            //Плановые часы нагрузки преподавателя
            double PlanHours = 0;            
            
            IList<clsDistribution> coll;

            //Начинаем глобальный обсчёт

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                else
                {
                    if (optCombine.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                    }
                    else
                    {
                        coll = mdlData.colDistribution;
                    }
                }
            }

            //Перебираем всех преподавателей
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Сбрасываем фактические часы нагрузки
                FactHours = 0;
                //Перебираем все строки распределения нагрузки
                for (int j = 0; j <= coll.Count - 1; j++)
                {
                    //Если строка участвует в расчёте нагрузки
                    if (!coll[j].flgExclude)
                    {
                        //Для случая обычных строк нагрузки
                        if (!coll[j].flgDistrib)
                        {
                            //Если нагрузка относится к какому-либо конкретному
                            //преподавателю
                            if (!(coll[j].Lecturer == null))
                            {
                                //Если выбранный преподаватель соответстует
                                //указанному в строке нагрузки
                                if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                                {
                                    //Суммируем часы нагрузки в этой строке
                                    FactHours += mdlData.toSumDistributionComponents(coll[i]);
                                }
                            }
                        }
                        //Для случаев равномерно распределяемой нагрузки
                        else
                        {
                            //Перебираем студентов
                            for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                if (mdlData.colStudents[k].flgPlan)
                                {
                                    //Если рассматриваемый преподаватель - руководитель студента
                                    //И если студент на том же курсе, где и дисциплина
                                    //И специальность студента должна соответствовать специальности нагрузки
                                    if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                        & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                        & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                    {
                                        FactHours += coll[j].Weight;
                                    }
                                }
                            }

                            //
                            if (optCombine.Checked)
                            {
                                mdlData.toDetectUniformInHoured(ref FactHours, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                            }
                        }
                    }
                }

                //Если у преподавателя нет разгрузки
                if (mdlData.colLecturer[i].UnLoad == 0)
                {
                    //Он должен отработать среднюю разгрузку других
                    UnLoadDopFact = mdlData.colLecturer[i].Rate * UnLoadDopValue;
                }
                else
                {
                    UnLoadDopFact = 0;
                }

                //Пересчитываем плановую нагрузку с учётом
                //средней разгрузки других
                if (optCombine.Checked || optMain.Checked)
                {
                    PlanHours = (mdlData.colLecturer[i].Rate * AvgLoad) +
                                 mdlData.colLecturer[i].UnLoad + UnLoadDopFact;
                }
                else
                {
                    PlanHours = 0;
                }

                Differ = FactHours - PlanHours;
                //Если у преподавателя получается перегрузка
                if (Differ > 0)
                {
                    //Суммируем перегрузку
                    PositiveSum += Differ;
                }

                //Если у преподавателя получается недогрузка
                if (Differ < 0)
                {
                    //Суммируем недогрузку
                    NegativeSum += Differ;
                }
            }

            //Суммарная перегрузка
            txtOverLoadSum.Text = PositiveSum.ToString();
            //Суммарный недогруз
            txtSumNonLoad.Text = NegativeSum.ToString();
            //Баланс между перегрузом и недогрузом
            txtDiffer.Text = (PositiveSum + NegativeSum).ToString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //Фамилия, имя, отчество
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO = (txtSurname.Text + " " + txtName.Text + " " +
                                                                     txtPatronymic.Text).Trim();
            //Должность
            if (cmbDutyList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty = mdlData.colDuty[cmbDutyList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty = null;
            }

            //Дополнительная должность
            if (cmbDutyAddList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1 = mdlData.colDuty[cmbDutyAddList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1 = null;
            }

            //Учёная степень
            if (cmbDegreeList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree = mdlData.colDegree[cmbDegreeList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree = null;
            }
            
            //Учёное звание
            if (cmbStatusList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status = mdlData.colStatus[cmbStatusList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status = null;
            }
            
            //Кафедра
            if (cmbDepartmentList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Depart = mdlData.colDepart[cmbDepartmentList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Depart = null;
            }

            //Дата рождения
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].DOB = Convert.ToDateTime(txtDOB.Text);
            
            //Стаж
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Seniority = Convert.ToDateTime(txtWorkSince.Text);
            
            //Разгрузка
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].UnLoad = Convert.ToInt32(txtUnLoad.Text);

            //Максимальная нагрузка
            if (!(txtMaxLoad.Text == ""))
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].MaxLoad = Convert.ToInt32(txtMaxLoad.Text);
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].MaxLoad = 0;
            }

            //Совместительство
            if (cmbCombinationList.SelectedIndex >= 0)
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination = mdlData.colCombination[cmbCombinationList.SelectedIndex];
            }
            else
            {
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination = null;
            }
            
            //Ставку
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate = Convert.ToDouble(txtRate.Text);

            //Старую ставку
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].OldRate = Convert.ToDouble(txtOldRate.Text);

            //Примечание для диспетчерской
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Text = txtText.Text;

            //Предпочтения по аудиториям
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Preferences = txtPreferences.Text;

            //Признак членства в учёном совете
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Soviet = chkSoviet.Checked;

            //Признак переменной ставки
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].ChangeRate = chkChangeRate.Checked;

            //Ставку первого семестра
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate1 = Convert.ToDouble(txtRate1.Text);

            //Ставку второго семестра
            mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate2 = Convert.ToDouble(txtRate2.Text);

            //Перезаполняем список преподавателей
            FillLecturerList();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            int DelElem = 0;
            //Если удаляется последний элемент списка, то его можно
            //просто удалить
            if (cmbLecturerList.SelectedIndex == cmbLecturerList.Items.Count - 1)
            {
                //Проходим всю нагрузку
                //и убираем удаляемого преподавателя отовсюду
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colHouredDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colHouredDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colCombineDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colCombineDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                //Удаляем преподавателя из коллекции преподавателей
                mdlData.colLecturer.RemoveAt(mdlData.colLecturer.Count - 1);
                //Удаляем преподавателя из списка преподавателей
                cmbLecturerList.Items.RemoveAt(cmbLecturerList.Items.Count - 1);
                //Переводим выбранный индекс в списке на позицию вверх
                cmbLecturerList.SelectedIndex = cmbLecturerList.Items.Count - 1;
            }

            //А если не последний, то нужно удалять аккуратно,
            //заменяя коды у всех элементов, оставшихся после
            else
            {
                //Запоминаем индекс удаляемого элемента
                DelElem = cmbLecturerList.SelectedIndex;
                //Проходим всю нагрузку и убираем
                //отовсюду удаляемого преподавателя
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colHouredDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colHouredDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colHouredDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colHouredDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colHouredDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (!(mdlData.colCombineDistribution[i].Lecturer == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer = null;
                        }
                    }

                    if (!(mdlData.colCombineDistribution[i].Lecturer2 == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer2.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer2 = null;
                        }
                    }

                    if (!(mdlData.colCombineDistribution[i].Lecturer3 == null))
                    {
                        if (mdlData.colCombineDistribution[i].Lecturer3.FIO.Equals(mdlData.colLecturer[DelElem].FIO))
                        {
                            mdlData.colCombineDistribution[i].Lecturer3 = null;
                        }
                    }
                }

                mdlData.colLecturer.RemoveAt(DelElem);
                cmbLecturerList.Items.RemoveAt(DelElem);

                for (int i = DelElem; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    mdlData.colLecturer[i].Code = mdlData.colLecturer[i].Code - 1;
                }

                cmbLecturerList.SelectedIndex = DelElem;
                FillLecturerList();
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //Создаём новый объект класса "Преподаватель"
            clsLecturer Lect = new clsLecturer();
            //Код назначаем на единицу больше, чем количество
            //элементов в коллекции
            Lect.Code = mdlData.colLecturer.Count + 1;
            //Формируем задел для нового (абстрактного) преподавателя
            Lect.FIO = "Иванов Иван Иванович";
            Lect.DOB = Convert.ToDateTime("01.01.1950");
            Lect.Seniority = DateTime.Now;
            Lect.MaxLoad = 0;
            Lect.OldRate = 0;
            Lect.Rate = 0;
            Lect.Text = "Новое пожелание по занятости";
            Lect.Preferences = "Аудитории ";
            Lect.UnLoad = 0;
            //Добавляем объект в коллекцию
            mdlData.colLecturer.Add(Lect);
            //Заносим объект в список
            cmbLecturerList.Items.Add(mdlData.colLecturer[mdlData.colLecturer.Count - 1].Code + ". " + mdlData.colLecturer[mdlData.colLecturer.Count - 1].FIO);
            //Переходим к новому элементу списка
            cmbLecturerList.SelectedIndex = cmbLecturerList.Items.Count - 1;
        }

        private void btnAnalysis_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {

            }

            MessageBox.Show("");
        }

        private void chkNonOverLoad_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
