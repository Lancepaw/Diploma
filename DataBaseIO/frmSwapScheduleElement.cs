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
    public partial class frmSwapScheduleElement : Form
    {
        public frmSwapScheduleElement()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            //Закрыть
            Dispose();
            //
            frmEditScheduleElement f = new frmEditScheduleElement();
            //
            f.Owner = Owner;
            //
            f.ShowDialog();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            //Заменить элемент расписания
            SwapElement();
            //Закрыть
            Dispose();
        }

        //Метод обмена элементов расписания
        private void SwapElement()
        {
            clsSchedule Sch;
            //Перебираем элементы расписания и отыскиваем указанный в экранной форме
            for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
            {
                Sch = mdlData.colSchedule[i];
                if (Sch.Lecturer.Equals(mdlData.colLecturer[cmbLecturer.SelectedIndex]))
                {
                    if (Sch.WeekDay.Equals(mdlData.colWeekDays[cmbWeekDay.SelectedIndex]))
                    {
                        if (Sch.Time.Equals(mdlData.colPairTime[cmbPairTime.SelectedIndex]))
                        {
                            if (Sch.Week.Equals(mdlData.colWeek[cmbWeek.SelectedIndex]))
                            {
                                if (Sch.Semestr.Equals(mdlData.colSemestr[cmbSemestr.SelectedIndex]))
                                {
                                    //из выбранного глобального элемента переписываем всё в найденный
                                    Sch.Subj = mdlData.SelectedScheduleElement.Subj;
                                    Sch.Auditory = mdlData.SelectedScheduleElement.Auditory;
                                    Sch.Group = mdlData.SelectedScheduleElement.Group;
                                    Sch.KursNum = mdlData.SelectedScheduleElement.KursNum;
                                    Sch.Link = mdlData.SelectedScheduleElement.Link;
                                    Sch.Spec = mdlData.SelectedScheduleElement.Spec;
                                    Sch.Stream = mdlData.SelectedScheduleElement.Stream;
                                    Sch.Subject = mdlData.SelectedScheduleElement.Subject;
                                    Sch.SubjectType = mdlData.SelectedScheduleElement.SubjectType;

                                    //выбранный элемент сбрасываем к умолчаниям
                                    mdlData.SelectedScheduleElement.Subj = false;
                                    mdlData.SelectedScheduleElement.Auditory = "";
                                    mdlData.SelectedScheduleElement.Group = "";
                                    mdlData.SelectedScheduleElement.KursNum = null;
                                    mdlData.SelectedScheduleElement.Link = null;
                                    mdlData.SelectedScheduleElement.Spec = null;
                                    mdlData.SelectedScheduleElement.Stream = "";
                                    mdlData.SelectedScheduleElement.Subject = null;
                                    mdlData.SelectedScheduleElement.SubjectType = null;

                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void frmSwapScheduleElement_Load(object sender, EventArgs e)
        {
            FillLecturerList();
            cmbLecturer.SelectedIndex = mdlData.colLecturer.IndexOf(mdlData.SelectedScheduleElement.Lecturer);

            FillWeekDayList();
            cmbWeekDay.SelectedIndex = mdlData.colWeekDays.IndexOf(mdlData.SelectedScheduleElement.WeekDay);

            FillPairTimeList();
            cmbPairTime.SelectedIndex = mdlData.colPairTime.IndexOf(mdlData.SelectedScheduleElement.Time);

            FillWeekList();
            cmbWeek.SelectedIndex = mdlData.colWeek.IndexOf(mdlData.SelectedScheduleElement.Week);

            FillSemestrList();
            cmbSemestr.SelectedIndex = mdlData.colSemestr.IndexOf(mdlData.SelectedScheduleElement.Semestr);
        }

        private void FillLecturerList()
        {
            //Очищаем перечень преподавателей
            cmbLecturer.Items.Clear();
            //Заполняем комбо-список преподавателями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturer.Items.Add(mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
            }
        }

        private void FillWeekDayList()
        {
            //Очищаем перечень дней недели
            cmbWeekDay.Items.Clear();
            //Заполняем комбо-список днями недели
            for (int i = 0; i <= mdlData.colWeekDays.Count - 1; i++)
            {
                cmbWeekDay.Items.Add(mdlData.colWeekDays[i].WeekDay);
            }
        }

        private void FillPairTimeList()
        {
            //Очищаем перечень времён пар
            cmbPairTime.Items.Clear();
            //Заполняем комбо-список временами пар
            for (int i = 0; i <= mdlData.colPairTime.Count - 1; i++)
            {
                cmbPairTime.Items.Add(mdlData.colPairTime[i].Time);
            }
        }

        private void FillWeekList()
        {
            //Очищаем перечень недель
            cmbWeek.Items.Clear();
            //Заполняем комбо-список неделями
            for (int i = 0; i <= mdlData.colWeek.Count - 1; i++)
            {
                cmbWeek.Items.Add(mdlData.colWeek[i].NumberWeek);
            }
        }

        private void FillSemestrList()
        {
            //Очищаем перечень семестров
            cmbSemestr.Items.Clear();
            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestr.Items.Add(mdlData.colSemestr[i].SemNum);
            }
        }
    }
}
