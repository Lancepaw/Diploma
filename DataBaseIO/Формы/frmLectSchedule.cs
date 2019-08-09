using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Globalization;

namespace DataBaseIO
{
    public partial class frmLectSchedule : Form
    {
        public frmLectSchedule()
        {
            InitializeComponent();            
        }

        private IList<clsSchedule> colScheduleProjection = new List<clsSchedule>();
        private int prLect = 0;
        private int prWeek = 0;
        private int prSem = 0;

        private void frmLectSchedule_Load(object sender, EventArgs e)
        {
            //Подключения актуального времени пар
            lblTime1.Text = mdlData.colPairTime[0].Time;
            lblTime2.Text = mdlData.colPairTime[1].Time;
            lblTime3.Text = mdlData.colPairTime[2].Time;
            //Четвёртая ячейка занята обедом
            lblTime5.Text = mdlData.colPairTime[3].Time;
            lblTime6.Text = mdlData.colPairTime[4].Time;
            lblTime7.Text = mdlData.colPairTime[5].Time;
            lblTime8.Text = mdlData.colPairTime[6].Time;
            lblTime9.Text = mdlData.colPairTime[7].Time;

            FillSemestrList();
            FillLecturerList();
            FillWeekList();

            TimeTable.DoubleClick += new EventHandler(TimeTable_DoubleClick);
            //Принудительное выставление интересующей даты
            dpStartSemestr.Value = Convert.ToDateTime("07/02/2018", new CultureInfo("ru-RU"));
        }

        private void TimeTable_DoubleClick(object sender, EventArgs e)
        {
            bool flgDetect;
            clsSchedule Sch;
            MouseEventArgs ms = (e as MouseEventArgs);
            mdlData.SelectedLecturer = mdlData.colLecturer[cmbLecturerList.SelectedIndex - 1];
            frmDistributionAccept fDA = new frmDistributionAccept();
            frmEditScheduleElement fESE = new frmEditScheduleElement();

            int[] widths = TimeTable.GetColumnWidths();
            int[] heights = TimeTable.GetRowHeights();

            int col = -1;
            int left = ms.X;
            for (int i = 0; i < widths.Length; i++)
            {
                if (left < widths[i])
                {
                    col = i;

                    break;
                }
                else
                {
                    left -= widths[i];
                }
            }

            int row = -1;
            int top = ms.Y;
            for (int i = 0; i < heights.Length; i++)
            {
                if (top < heights[i])
                {
                    row = i;
                    break;
                }
                else
                {
                    top -= heights[i];
                }
            }

            //Если не заголовки таблицы расписания и не
            //4-й столбец обеденного перерыва
            if (row > 0 & col > 0 & col != 4)
            {
                //Если попадаем в область правее обеденного перерыва,
                //необходимо уменьшить индекс столбца на единицу
                if (col > 4)
                {
                    col--;
                }

                MessageBox.Show("Строка: " + mdlData.colWeekDays[row - 1].WeekDay + 
                                "; Столбец: " + mdlData.colPairTime[col - 1].Time);

                //Сбрасываем флаг обнаружения
                flgDetect = false;
                //Перебираем элементы расписания
                for (int i = 0; i < mdlData.colSchedule.Count; i++)
                {
                    //Фиксируем i-й элемент расписания
                    Sch = mdlData.colSchedule[i];
                    //Если совпали преподаватели
                    if (Sch.Lecturer.FIO.Equals(mdlData.SelectedLecturer.FIO))
                    {
                        //Если совпали семестры
                        if (Sch.Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemestr.SelectedIndex].SemNum))
                        {
                            //Если совпали учебные недели
                            if (Sch.Week.NumberWeek.Equals(mdlData.colWeek[cmbWeekList.SelectedIndex - 1].NumberWeek))
                            {
                                //Если совпали дни недели
                                if (Sch.WeekDay.WeekDay.Equals(mdlData.colWeekDays[row - 1].WeekDay))
                                {
                                    //Если совпали дни недели
                                    if (Sch.Time.Time.Equals(mdlData.colPairTime[col - 1].Time))
                                    {
                                        //Если для элемента расписания выставлен признак наличия занятия
                                        if (Sch.Subj)
                                        {
                                            //Выставляем флаг обнаружения
                                            flgDetect = true;
                                            //Записываем в глобальную переменную
                                            //найденный элемент расписания
                                            mdlData.SelectedScheduleElement = Sch;
                                            //прерываем цикл
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //Если элемент расписания обнаружен
                if (flgDetect)
                {
                    //fDA.ShowDialog(this);
                    //Открываем окно редактирования элемента расписания
                    fESE.ShowDialog(this);
                }
            }
        }

        //Заполняем перечень преподавателей
        private void FillLecturerList()
        {
            //Очищаем комбо-список преподавателей
            cmbLecturerList.Items.Clear();
            //Добавляем нулевой элемент - не выбранный
            cmbLecturerList.Items.Add("не выбран");
            
            //Заполняем комбо-список преподавателями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturerList.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }
        }

        //Заполняем перечень недель
        private void FillWeekList()
        {
            //Очищаем список
            cmbWeekList.Items.Clear();

            cmbWeekList.Items.Add("не выбрана");

            //Заполняем комбо-список неделями
            for (int i = 0; i <= mdlData.colWeek.Count - 1; i++)
            {
                cmbWeekList.Items.Add(mdlData.colWeek[i].NumberWeek);
            }
        }

        //Заполняем перечень семестров
        private void FillSemestrList()
        {
            //Очищаем комбо-список
            cmbSemestr.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestr.Items.Add(mdlData.colSemestr[i].SemNum);
            }
        }

        private void cmbWeekList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetAllFlags();
            SetProjection(cmbLecturerList.SelectedIndex);
            SetTimes(cmbWeekList.SelectedIndex, cmbSemestr.SelectedIndex);
            prWeek = cmbWeekList.SelectedIndex;
        }

        private void btnSaveSchedule_Click(object sender, EventArgs e)
        {
            int Lect = cmbLecturerList.SelectedIndex;
            int Week = cmbWeekList.SelectedIndex;
            int Semestr = cmbSemestr.SelectedIndex;

            if (Lect > 0 & Week > 0 & Semestr > 0)
            {
                for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
                {
                    if (mdlData.colSchedule[i].Lecturer.Equals(mdlData.colLecturer[Lect - 1]) &
                        mdlData.colSchedule[i].Week.Equals(mdlData.colWeek[Week - 1]) &
                        mdlData.colSchedule[i].Semestr.Equals(mdlData.colSemestr[Semestr]))
                    {
                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Понедельник"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkMonday2000.Checked;
                            }
                        }

                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Вторник"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkTuesday2000.Checked;
                            }
                        }

                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Среда"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkWednesday2000.Checked;
                            }
                        }

                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Четверг"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkThursday2000.Checked;
                            }
                        }

                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Пятница"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkFriday2000.Checked;
                            }
                        }

                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Суббота"))
                        {
                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[0].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday800.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[1].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday940.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[2].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday1120.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[3].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday1320.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[4].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday1500.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[5].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday1640.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[6].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday1820.Checked;
                            }

                            if (mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[7].Time))
                            {
                                mdlData.colSchedule[i].Subj = chkSaturday2000.Checked;
                            }
                        }
                    }
                }

                cmbLecturerList.SelectedIndex = prLect;
                cmbWeekList.SelectedIndex = prWeek;
                cmbSemestr.SelectedIndex = prSem;
            }
            else
            {
                //Оповещение
                MessageBox.Show("Невозможно выполнить сохранение.\nНедостаточно данных", "Ошибка", MessageBoxButtons.OK);
            }
        }

        private void cmbLecturerList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetAllFlags();
            SetProjection(cmbLecturerList.SelectedIndex);
            SetTimes(cmbWeekList.SelectedIndex, cmbSemestr.SelectedIndex);
            prLect = cmbLecturerList.SelectedIndex;
        }

        //Составляем проекцию элементов расписания строго под преподавателя
        private void SetProjection(int Lect)
        {
            clsSchedule Sc;
            colScheduleProjection.Clear();

            //Если преподаватель указан
            if (Lect > 0)
            {
                //Перебираем элементы расписания
                for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
                {
                    Sc = mdlData.colSchedule[i];
                    //Если преподаватель в элементе расписания совпал с указанным в списке
                    if (Sc.Lecturer.Equals(mdlData.colLecturer[Lect - 1]))
                    {
                        //И если он что-то ведёт у студентов, согласно
                        //выбранному элементу расписания, то
                        if (Sc.Subj)
                        {
                            //добавляем этот элемент в проекцию
                            colScheduleProjection.Add(Sc);
                        }
                    }
                }
            }
        }

        private void SetTimes(int Week, int Semestr)
        {
            clsSchedule Sc;

            //Если создана проекция расписания, если выбрана неделя и выбран семестр
            if (colScheduleProjection.Count > 0 & Week > 0 & Semestr > 0)
            {
                //Перебираем все элементы расписания в проекции
                for (int i = 0; i <= colScheduleProjection.Count - 1; i++)
                {
                    Sc = colScheduleProjection[i];

                    //Если семестр именно тот, что выбран в списке
                    if (Sc.Semestr.Equals(mdlData.colSemestr[Semestr]))
                    {
                        //Если неделя именно та, которая выбрана в списке
                        if (Sc.Week.Equals(mdlData.colWeek[Week - 1]))
                        {
                            if (Sc.WeekDay.WeekDay.Equals("Понедельник"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkMonday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkMonday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkMonday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkMonday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkMonday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkMonday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkMonday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkMonday2000.Checked = Sc.Subj;
                                }
                            }

                            if (Sc.WeekDay.WeekDay.Equals("Вторник"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkTuesday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkTuesday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkTuesday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkTuesday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkTuesday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkTuesday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkTuesday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkTuesday2000.Checked = Sc.Subj;
                                }
                            }

                            if (Sc.WeekDay.WeekDay.Equals("Среда"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkWednesday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkWednesday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkWednesday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkWednesday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkWednesday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkWednesday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkWednesday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkWednesday2000.Checked = Sc.Subj;
                                }
                            }

                            if (Sc.WeekDay.WeekDay.Equals("Четверг"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkThursday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkThursday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkThursday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkThursday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkThursday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkThursday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkThursday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkThursday2000.Checked = Sc.Subj;
                                }
                            }

                            if (Sc.WeekDay.WeekDay.Equals("Пятница"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkFriday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkFriday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkFriday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkFriday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkFriday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkFriday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkFriday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkFriday2000.Checked = Sc.Subj;
                                }
                            }

                            if (Sc.WeekDay.WeekDay.Equals("Суббота"))
                            {
                                if (Sc.Time.Time.Equals(mdlData.colPairTime[0].Time))
                                {
                                    chkSaturday800.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[1].Time))
                                {
                                    chkSaturday940.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[2].Time))
                                {
                                    chkSaturday1120.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[3].Time))
                                {
                                    chkSaturday1320.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[4].Time))
                                {
                                    chkSaturday1500.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[5].Time))
                                {
                                    chkSaturday1640.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[6].Time))
                                {
                                    chkSaturday1820.Checked = Sc.Subj;
                                }

                                if (Sc.Time.Time.Equals(mdlData.colPairTime[7].Time))
                                {
                                    chkSaturday2000.Checked = Sc.Subj;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                ResetAllFlags();
            }
        }

        private void ResetAllFlags()
        {
            //Сбросить все флажки
            chkFriday1120.Checked = false;
            chkFriday1320.Checked = false;
            chkFriday1500.Checked = false;
            chkFriday1640.Checked = false;
            chkFriday1820.Checked = false;
            chkFriday2000.Checked = false;
            chkFriday800.Checked = false;
            chkFriday940.Checked = false;

            chkMonday1120.Checked = false;
            chkMonday1320.Checked = false;
            chkMonday1500.Checked = false;
            chkMonday1640.Checked = false;
            chkMonday1820.Checked = false;
            chkMonday2000.Checked = false;
            chkMonday800.Checked = false;
            chkMonday940.Checked = false;

            chkSaturday1120.Checked = false;
            chkSaturday1320.Checked = false;
            chkSaturday1500.Checked = false;
            chkSaturday1640.Checked = false;
            chkSaturday1820.Checked = false;
            chkSaturday2000.Checked = false;
            chkSaturday800.Checked = false;
            chkSaturday940.Checked = false;

            chkThursday1120.Checked = false;
            chkThursday1320.Checked = false;
            chkThursday1500.Checked = false;
            chkThursday1640.Checked = false;
            chkThursday1820.Checked = false;
            chkThursday2000.Checked = false;
            chkThursday800.Checked = false;
            chkThursday940.Checked = false;

            chkTuesday1120.Checked = false;
            chkTuesday1320.Checked = false;
            chkTuesday1500.Checked = false;
            chkTuesday1640.Checked = false;
            chkTuesday1820.Checked = false;
            chkTuesday2000.Checked = false;
            chkTuesday800.Checked = false;
            chkTuesday940.Checked = false;

            chkWednesday1120.Checked = false;
            chkWednesday1320.Checked = false;
            chkWednesday1500.Checked = false;
            chkWednesday1640.Checked = false;
            chkWednesday1820.Checked = false;
            chkWednesday2000.Checked = false;
            chkWednesday800.Checked = false;
            chkWednesday940.Checked = false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            int Lect = cmbLecturerList.SelectedIndex;
            int Week = cmbWeekList.SelectedIndex;

            lstSwap.Items.Clear();

            if (Lect > 0 & Week > 0)
            {
                for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
                {
                    if (!mdlData.colSchedule[i].Lecturer.FIO.Equals(mdlData.colLecturer[Lect - 1].FIO))
                    {
                        if (mdlData.colSchedule[i].Week.Equals(mdlData.colWeek[Week - 1]))
                        {
                            for (int j = 0; j <= mdlData.colSchedule.Count - 1; j++)
                            {
                                if (mdlData.colSchedule[j].Lecturer.FIO.Equals(mdlData.colLecturer[Lect - 1].FIO))
                                {
                                    if (mdlData.colSchedule[j].Week.Equals(mdlData.colWeek[Week - 1]))
                                    {
                                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals("Вторник") &
                                            mdlData.colSchedule[j].WeekDay.WeekDay.Equals("Вторник"))
                                        {
                                            if (mdlData.colSchedule[i].Time.Time.Equals("13.20-14.50") &
                                                mdlData.colSchedule[j].Time.Time.Equals("13.20-14.50"))
                                            {
                                                if (mdlData.colSchedule[i].Subj & mdlData.colSchedule[j].Subj)
                                                {
                                                    lstSwap.Items.Add(mdlData.colSchedule[i].Lecturer.FIO);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void btnSwap_Click(object sender, EventArgs e)
        {
            int Lect = cmbLecturerList.SelectedIndex;
            int Week = cmbWeekList.SelectedIndex;
            int unitCountIll = 0;
            int unitCountSwap = 0;

            bool flg = false;

            lstSwap.Items.Clear();

            //Если выбран какой-то преподаватель и какая-то неделя
            if (Lect > 0 & Week > 0)
            {
                //Начинаем перебирать все элементы расписания
                for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
                {
                    //Если i элемент расписания принадлежит преподавателю, который
                    //выбран в списке, то элемент подходит
                    if (mdlData.colSchedule[i].Lecturer.FIO.Equals(mdlData.colLecturer[Lect - 1].FIO))
                    {
                        //Если i элемент расписания принадлежит неделе, которая
                        //выбрана в списке, то элемент подходит                        
                        if (mdlData.colSchedule[i].Week.Equals(mdlData.colWeek[Week - 1]))
                        {
                            //Перебираем дни недели
                            for (int d = 0; d <= mdlData.colWeekDays.Count - 1; d++)
                            {
                                //Перебираем времена пар
                                for (int t = 0; t <= mdlData.colPairTime.Count - 1; t++)
                                {
                                    //Если в рассматреваемый день и рассматриваемое время
                                    //элемент расписания существует
                                    if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals(mdlData.colWeekDays[d].WeekDay) &
                                        mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[t].Time))
                                    {
                                        //(галка на пересечении выставлена)
                                        if (mdlData.colSchedule[i].Subj)
                                        {
                                            //увеличиваем количество пар, которые проводит преподаватель
                                            unitCountIll++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //Пишем, сколько занятий за выбранную неделю проводит выбранный преподаватель
                lstSwap.Items.Add(mdlData.colLecturer[Lect - 1].FIO + " ведёт " + unitCountIll + " занятий");

                //Начинаем перебирать преподавателей кафедры
                //Всем, кроме выбранного из списка, будет соответствовать
                //индекс j
                for (int j = 0; j <= mdlData.colLecturer.Count - 1; j++)
                {
                    //сбрасываем счётчик возможных замен заменяющим заболевшего
                    unitCountSwap = 0;
                    //Если j преподатель не совпадает с тем, который выбран
                    //в списке, то он нам подходит
                    if (!mdlData.colLecturer[j].FIO.Equals(mdlData.colLecturer[Lect - 1].FIO))
                    {
                        //Начинаем искать элементы расписания для выбранного преподавателя
                        //им соответствует индекс i
                        for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
                        {
                            //Если i элемент расписания принадлежит преподавателю, который
                            //выбран в списке, то элемент подходит
                            if (mdlData.colSchedule[i].Lecturer.FIO.Equals(mdlData.colLecturer[Lect - 1].FIO))
                            {
                                //Если i элемент расписания принадлежит неделе, которая
                                //выбрана в списке, то элемент подходит                        
                                if (mdlData.colSchedule[i].Week.Equals(mdlData.colWeek[Week - 1]))
                                {
                                    //Начинаем искать элементы расписания для j преподавателя
                                    //им соответствует индекс k
                                    for (int k = 0; k <= mdlData.colSchedule.Count - 1; k++)
                                    {
                                        //Если k элемент расписания принадлежит преподавателю j,
                                        //то элемент подходит
                                        if (mdlData.colSchedule[k].Lecturer.FIO.Equals(mdlData.colLecturer[j].FIO))
                                        {
                                            //Если k элемент расписания принадлежит неделе, которая
                                            //выбрана в списке, то элемент подходит                                             
                                            if (mdlData.colSchedule[k].Week.Equals(mdlData.colWeek[Week - 1]))
                                            {
                                                //Перебираем дни недели
                                                for (int d = 0; d <= mdlData.colWeekDays.Count - 1; d++)
                                                {
                                                    //Перебираем времена пар
                                                    for (int t = 0; t <= mdlData.colPairTime.Count - 1; t++)
                                                    {
                                                        //Если тот и другой элементы расписания совпадают по
                                                        //времени и дню недели
                                                        if (mdlData.colSchedule[i].WeekDay.WeekDay.Equals(mdlData.colWeekDays[d].WeekDay) &
                                                            mdlData.colSchedule[k].WeekDay.WeekDay.Equals(mdlData.colWeekDays[d].WeekDay) &
                                                            mdlData.colSchedule[i].Time.Time.Equals(mdlData.colPairTime[t].Time) &
                                                            mdlData.colSchedule[k].Time.Time.Equals(mdlData.colPairTime[t].Time))
                                                        {
                                                            //Если элемент выбранного из списка преподавателя в это время
                                                            //существует (галка), а элемент j преподавателя - не существует
                                                            //(галки нет)
                                                            if (mdlData.colSchedule[i].Subj & !mdlData.colSchedule[k].Subj)
                                                            {
                                                                //увеличиваем счётчик возможных замен заменяющим заболевшего
                                                                //на единицу
                                                                unitCountSwap++;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                   
                    //Если преподаватель j свободен во все часы, которые занят
                    //выбранный из списка преподаватель, то он подходит
                    if (unitCountIll == unitCountSwap)
                    {
                        //Заменять друг друга могут преподаватели одного направления
                        //(пока это разделение по кафедре)
                        if (mdlData.colLecturer[j].Depart.Short.Equals(mdlData.colLecturer[Lect - 1].Depart.Short))
                        {
                            //по умолчанию считаем, что преподаватель j ничего не ведёт
                            flg = false;
                            //ищем нагрузку преподавателя j
                            for (int ds = 0; ds <= mdlData.colCombineDistribution.Count - 1; ds++)
                            {
                                //Если в перечне нагрузки найден j преподаватель
                                if (mdlData.colCombineDistribution[ds].Lecturer.FIO.Equals(mdlData.colLecturer[j].FIO))
                                {
                                    //Если j преподаватель ведёт лекции или принимает экзамены,
                                    //зачёты, рефераты, проводит консультации или в ответе
                                    //за какой-то вид нагрузки (часы не нулевые),
                                    //то считается, что от работает в штате
                                    if (mdlData.colCombineDistribution[ds].Lecture > 0 ||
                                        mdlData.colCombineDistribution[ds].Exam > 0 ||
                                        mdlData.colCombineDistribution[ds].Credit > 0 ||
                                        mdlData.colCombineDistribution[ds].RefHomeWork > 0 ||
                                        mdlData.colCombineDistribution[ds].Tutorial > 0 ||
                                        mdlData.colCombineDistribution[ds].LabWork > 0 ||
                                        mdlData.colCombineDistribution[ds].Practice > 0 ||
                                        mdlData.colCombineDistribution[ds].IndividualWork > 0 ||
                                        mdlData.colCombineDistribution[ds].KRAPK > 0 ||
                                        mdlData.colCombineDistribution[ds].KursProject > 0 ||
                                        mdlData.colCombineDistribution[ds].PreDiplomaPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].DiplomaPaper > 0 ||
                                        mdlData.colCombineDistribution[ds].TutorialPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].ProducingPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].GAK > 0 ||
                                        mdlData.colCombineDistribution[ds].PostGrad > 0 ||
                                        mdlData.colCombineDistribution[ds].Visiting > 0 ||
                                        mdlData.colCombineDistribution[ds].Magistry > 0)
                                    {
                                        //выставляем признак работающего в штате преподавателя
                                        flg = true;
                                        //если хоть что-то найдено, то далее можно не искать
                                        break;
                                    }
                                }
                            }

                            //Если преподаватель j хоть что-то ведёт, то он
                            //может быть заменой
                            if (flg)
                            {
                                //записываем подходящего в список
                                lstSwap.Items.Add("Полностью может заменить: " + mdlData.colLecturer[j].FIO);
                            }
                        }
                    }
                    //Если возможность замены полностью не покрывает
                    //расписание заболевшего, то это частичная замена
                    else
                    {
                        //Заменять друг друга могут преподаватели одного направления
                        //(пока это разделение по кафедре)
                        if (mdlData.colLecturer[j].Depart.Short.Equals(mdlData.colLecturer[Lect - 1].Depart.Short))
                        {
                            //по умолчанию считаем, что преподаватель j ничего не ведёт
                            flg = false;
                            //ищем нагрузку преподавателя j
                            for (int ds = 0; ds <= mdlData.colCombineDistribution.Count - 1; ds++)
                            {
                                //Если в перечне нагрузки найден j преподаватель
                                if (mdlData.colCombineDistribution[ds].Lecturer.FIO.Equals(mdlData.colLecturer[j].FIO))
                                {
                                    //Если j преподаватель ведёт лекции или принимает экзамены,
                                    //зачёты, рефераты, проводит консультации или в ответе
                                    //за какой-то вид нагрузки (часы не нулевые),
                                    //то считается, что от работает в штате
                                    if (mdlData.colCombineDistribution[ds].Lecture > 0 ||
                                        mdlData.colCombineDistribution[ds].Exam > 0 ||
                                        mdlData.colCombineDistribution[ds].Credit > 0 ||
                                        mdlData.colCombineDistribution[ds].RefHomeWork > 0 ||
                                        mdlData.colCombineDistribution[ds].Tutorial > 0 ||
                                        mdlData.colCombineDistribution[ds].LabWork > 0 ||
                                        mdlData.colCombineDistribution[ds].Practice > 0 ||
                                        mdlData.colCombineDistribution[ds].IndividualWork > 0 ||
                                        mdlData.colCombineDistribution[ds].KRAPK > 0 ||
                                        mdlData.colCombineDistribution[ds].KursProject > 0 ||
                                        mdlData.colCombineDistribution[ds].PreDiplomaPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].DiplomaPaper > 0 ||
                                        mdlData.colCombineDistribution[ds].TutorialPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].ProducingPractice > 0 ||
                                        mdlData.colCombineDistribution[ds].GAK > 0 ||
                                        mdlData.colCombineDistribution[ds].PostGrad > 0 ||
                                        mdlData.colCombineDistribution[ds].Visiting > 0 ||
                                        mdlData.colCombineDistribution[ds].Magistry > 0)
                                    {
                                        //выставляем признак работающего в штате преподавателя
                                        flg = true;
                                        //если хоть что-то найдено, то далее можно не искать
                                        break;
                                    }
                                }
                            }

                            //Если частичная замена возможна самим же собой,
                            //то принудительно исключить из рассмотрения
                            if (j == Lect - 1)
                            {
                                flg = false;
                            }

                            //Если преподаватель с частичной заменой по условиям
                            //проходит, то добавляем его в коллекцию
                            if (flg)
                            {
                                clsLecturer_Load LL = new clsLecturer_Load();
                                LL.Code = mdlData.colLectLoad.Count + 1;
                                LL.Lecturer = mdlData.colLecturer[j];
                                LL.Load = unitCountSwap;
                                mdlData.colLectLoad.Add(LL);
                            }
                        }
                    }
                }

                //Упорядочивание по убыванию возможности замены
                //заменяющими преподавателями
                for (int i = 0; i <= mdlData.colLectLoad.Count - 2; i++)
                {
                    for (int j = i + 1; j <= mdlData.colLectLoad.Count - 1; j++)
                    {
                        //Если впереди (справа от текущего i) элемент с большими возможностями,
                        //его необходимо поднять по списку (повысить приоритет)
                        if (mdlData.colLectLoad[i].Load < mdlData.colLectLoad[j].Load)
                        {
                            //Промежуточный объект-переменная
                            clsLecturer_Load LLtmp = new clsLecturer_Load();
                            LLtmp = mdlData.colLectLoad[i];
                            mdlData.colLectLoad[i] = mdlData.colLectLoad[j];
                            mdlData.colLectLoad[j] = LLtmp;
                            LLtmp = null;
                        }
                    }
                }

                //Дополняем список заменяющих теми, кто может это сделать частично
                for (int i = 0; i <= mdlData.colLectLoad.Count - 1; i++)
                {
                    lstSwap.Items.Add("Частично может заменить: " + 
                        mdlData.colLectLoad[i].Lecturer.FIO + " - " + 
                        Convert.ToInt32((Convert.ToDouble(mdlData.colLectLoad[i].Load) / Convert.ToDouble(unitCountIll)) * 100) + "%");
                }
            }
        }

        private void btnRealWork_Click(object sender, EventArgs e)
        {
            int workCounter = 0;
            int sickCounter = 0;

            bool flgSick;

            bool flgFirstWeek = true;

            string outString = "";

            //Устанавливаем текущей датой дату начала семестра
            DateTime currDate = dpStartSemestr.Value;

            clsWeekDays WD = null;
            clsSchedule Sc = null;
            clsWeek W = null;
            clsScheduleSickHours SSH = null;
            clsScheduleSickHours SSH_Same = null;

            //Очищаем коллекцию
            mdlData.colScheduleSickHours.Clear();

            //Перебираем календарные дни с начала семестра до конца семестра
            for (int i = 1; i <= 18 * 7; i++)
            {
                switch (currDate.DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("понедельник"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }

                            if (currDate != dpStartSemestr.Value)
                            {
                                if (flgFirstWeek)
                                {
                                    flgFirstWeek = false;
                                }
                                else
                                {
                                    flgFirstWeek = true;
                                }
                            }

                            break;
                        }

                    case DayOfWeek.Tuesday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("вторник"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    case DayOfWeek.Wednesday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("среда"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    case DayOfWeek.Thursday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("четверг"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    case DayOfWeek.Friday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("пятница"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    case DayOfWeek.Saturday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("суббота"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    case DayOfWeek.Sunday:
                        {
                            for (int j = 0; j < mdlData.colWeekDays.Count; j++)
                            {
                                if (mdlData.colWeekDays[j].WeekDay.ToLower().Equals("воскресенье"))
                                {
                                    WD = mdlData.colWeekDays[j];
                                    break;
                                }
                            }
                            break;
                        }

                    default:
                        {
                            WD = null;
                            break;
                        }
                }
                
                if (flgFirstWeek)
                {
                    W = mdlData.colWeek[0];
                }
                else
                {
                    W = mdlData.colWeek[1];
                }

                flgSick = false;
                //Просматриваем больничные листы
                for (int j = 0; j < mdlData.colSickList.Count; j++)
                {
                    if (mdlData.colSickList[j].Lecturer.Equals(colScheduleProjection[0].Lecturer))
                    {
                        if (currDate >= mdlData.colSickList[j].OpenDate & currDate <= mdlData.colSickList[j].CloseDate)
                        {
                            flgSick = true;
                            break;
                        }
                    }
                }

                //Если преподаватель не болел в рассматриваемую дату
                if (!flgSick)
                {
                    //Перебираем элементы расписания
                    for (int j = 0; j < colScheduleProjection.Count; j++)
                    {
                        Sc = colScheduleProjection[j];

                        //Если день недели совпал и неделя совпала
                        if (Sc.WeekDay.Equals(WD) & Sc.Week.Equals(W))
                        {
                            //Перебираем времена
                            for (int k = 0; k < mdlData.colPairTime.Count; k++)
                            {
                                if (Sc.Time.Equals(mdlData.colPairTime[k]))
                                {
                                    if (Sc.Subj)
                                    {
                                        workCounter += 2;
                                    }
                                }
                            }
                        }
                    }
                }
                //Если преподаватель болел в рассматриваемую дату
                else
                {
                    //Перебираем элементы расписания
                    for (int j = 0; j < colScheduleProjection.Count; j++)
                    {
                        //Фиксируем текущий элемент расписания
                        Sc = colScheduleProjection[j];

                        //Если день недели совпал и неделя совпала
                        if (Sc.WeekDay.Equals(WD) & Sc.Week.Equals(W))
                        {
                            //Перебираем времена
                            for (int k = 0; k < mdlData.colPairTime.Count; k++)
                            {
                                //Если время занятия совпало
                                if (Sc.Time.Equals(mdlData.colPairTime[k]))
                                {
                                    //Если выставлен признак наличия занятия
                                    if (Sc.Subj)
                                    {
                                        SSH = new clsScheduleSickHours();
                                        SSH.Code = mdlData.colScheduleSickHours.Count + 1;
                                        SSH.Subject = Sc.Subject;
                                        SSH.SubjType = Sc.SubjectType;
                                        SSH.Spec = Sc.Spec;
                                        SSH.KursNum = Sc.KursNum;
                                        SSH.SickHours = 2;

                                        SSH_Same = null;
                                        for (int l = 0; l < mdlData.colScheduleSickHours.Count; l++)
                                        {
                                            if (mdlData.colScheduleSickHours[l].Subject != null &
                                                mdlData.colScheduleSickHours[l].SubjType != null &
                                                mdlData.colScheduleSickHours[l].KursNum != null &
                                                mdlData.colScheduleSickHours[l].Spec != null)
                                            {
                                                if (mdlData.colScheduleSickHours[l].Subject.Equals(SSH.Subject) &
                                                    mdlData.colScheduleSickHours[l].SubjType.Equals(SSH.SubjType) &
                                                    mdlData.colScheduleSickHours[l].KursNum.Equals(SSH.KursNum) &
                                                    mdlData.colScheduleSickHours[l].Spec.Equals(SSH.Spec))
                                                {
                                                    SSH_Same = mdlData.colScheduleSickHours[l];
                                                    SSH_Same.SickHours += 2;
                                                    break;
                                                }
                                            }
                                        }

                                        if (SSH_Same == null)
                                        {
                                            mdlData.colScheduleSickHours.Add(SSH);
                                        }

                                        sickCounter += 2;
                                    }
                                }
                            }
                        }
                    }
                }

                currDate = currDate.AddDays(1);
            }

            for (int i = 0; i < mdlData.colScheduleSickHours.Count; i++)
            {
                SSH = mdlData.colScheduleSickHours[i];
                if (SSH.Spec != null & SSH.KursNum != null & SSH.SubjType != null)
                {
                    outString += "\n" + SSH.Spec.ShortInstitute + "-" + SSH.KursNum.Kurs + " (" +
                        SSH.SubjType.Short + "): " + SSH.SickHours;
                }
                else
                {
                    outString += "\nНе определено: " + SSH.SickHours;
                }
            }

            MessageBox.Show("Всего проведено аудиторных занятий: " + workCounter +
                "\nНе проведено аудиторных занятий: " + sickCounter + outString);
        }

        private void dpStartSemestr_ValueChanged(object sender, EventArgs e)
        {

            //Известно, что учебный семестр длится 18 недель
            //С учётом того, что можно прибавлять только дни, месяцы, годы
            //пересчитаем недели через дни: 18 * 7 = 126 дней в семестре
            dpEndSemestr.Value = dpStartSemestr.Value.AddDays(18 * 7);
        }

        private void btnFindDoubles_Click(object sender, EventArgs e)
        {
            bool flgExecute = false;
            int counter = 0;

            //Находим и удаляем повторы
            for (int i = 0; i < mdlData.colSchedule.Count; i++)
            {
                for (int j = i + 1; j < mdlData.colSchedule.Count; j++)
                {
                    if (mdlData.colSchedule[i].Lecturer.Equals(mdlData.colSchedule[j].Lecturer) &
                        mdlData.colSchedule[i].Semestr.Equals(mdlData.colSchedule[j].Semestr) &
                        mdlData.colSchedule[i].Week.Equals(mdlData.colSchedule[j].Week) &
                        mdlData.colSchedule[i].WeekDay.Equals(mdlData.colSchedule[j].WeekDay) &
                        mdlData.colSchedule[i].Time.Equals(mdlData.colSchedule[j].Time))
                    {
                        mdlData.colSchedule.RemoveAt(j);
                        flgExecute = true;
                        counter++;
                        j--;
                    }
                }
            }

            //Выправляем коды элементов
            for (int i = 0; i < mdlData.colSchedule.Count; i++)
            {
                mdlData.colSchedule[i].Code = i + 1;
            }

            if (flgExecute)
            {
                MessageBox.Show("Удалено " + counter + " элементов расписания");
            }
            else
            {
                MessageBox.Show("Повторы не найдены");
            }
        }

        //Выбор семестра
        private void cmbSemestr_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetAllFlags();
            SetProjection(cmbLecturerList.SelectedIndex);
            SetTimes(cmbWeekList.SelectedIndex, cmbSemestr.SelectedIndex);
            prSem = cmbSemestr.SelectedIndex;
        }

        //Передача расписания другому преподавателю
        private void btnChange_Click(object sender, EventArgs e)
        {
            //Выбрать преподавателя, которому передаётся расписание
            mdlData.SelectedLecturer = null;
            mdlData.toGenerateForm(this, new frmLectInput());

            //Пройти по всем элементам расписания
            for (int i = 0; i < mdlData.colSchedule.Count; i++)
            {
                //Если выбранный для передачи преподаватель совпал с
                //текущим рассматриваемым, то продолжаем работу
                if (mdlData.SelectedLecturer.FIO.Equals(mdlData.colSchedule[i].Lecturer.FIO))
                {
                    //Пройти по всем элементам проекции расписания
                    //для выбранного преподавателя
                    for (int j = 0; j < colScheduleProjection.Count; j++)
                    {
                        if (colScheduleProjection[j].Semestr.Equals(mdlData.colSchedule[i].Semestr) &
                            colScheduleProjection[j].Week.Equals(mdlData.colSchedule[i].Week) &
                            colScheduleProjection[j].WeekDay.Equals(mdlData.colSchedule[i].WeekDay) &
                            colScheduleProjection[j].Time.Equals(mdlData.colSchedule[i].Time))
                        {
                            //Сначала переписываем выбранному преподавателю параметры найденного элемента
                            mdlData.colSchedule[i].Subj = colScheduleProjection[j].Subj;
                            mdlData.colSchedule[i].Auditory = colScheduleProjection[j].Auditory;
                            mdlData.colSchedule[i].KursNum = colScheduleProjection[j].KursNum;
                            mdlData.colSchedule[i].Subject = colScheduleProjection[j].Subject;
                            mdlData.colSchedule[i].SubjectType = colScheduleProjection[j].SubjectType;
                            mdlData.colSchedule[i].Spec = colScheduleProjection[j].Spec;

                            //Затем очищаем найденный элемент
                            colScheduleProjection[j].Subj = false;
                            colScheduleProjection[j].Auditory = "";
                            colScheduleProjection[j].KursNum = null;
                            colScheduleProjection[j].Subject = null;
                            colScheduleProjection[j].SubjectType = null;
                            colScheduleProjection[j].Spec = null;
                        }
                    }
                }
            }

            MessageBox.Show(this, "Выполнено");
        }

        private void btnCorrectGroups_Click(object sender, EventArgs e)
        {
            //Перебираем все элементы расписания
            for (int i = 0; i <= mdlData.colSchedule.Count - 1; i++)
            {
                if (mdlData.colSchedule[i].Group.Length > 1)
                {
                    mdlData.colSchedule[i].Group = mdlData.colSchedule[i].Group.Substring(mdlData.colSchedule[i].Group.Length - 1);
                }
            }

            MessageBox.Show(this, "Выполнено");
        }

        private void TimeTable_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
