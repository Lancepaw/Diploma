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
    public partial class frmEditScheduleElement : Form
    {
        public frmEditScheduleElement()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmEditScheduleElement_Load(object sender, EventArgs e)
        {
            FillSubjectList();
            FillSubjectTypesList();
            FillSpecialisationList();
            FillKursNumList();

            if (mdlData.SelectedScheduleElement != null)
            {
                //Если дисциплина указана
                if (mdlData.SelectedScheduleElement.Subject != null)
                {
                    //Если количество элементов в списке больше нуля
                    if (cmbSubject.Items.Count > 0)
                    {
                        cmbSubject.SelectedIndex = mdlData.colSubject.IndexOf(mdlData.SelectedScheduleElement.Subject);
                    }
                }
                else
                {
                    cmbSubject.SelectedIndex = -1;
                }

                //Если вид занятия указан
                if (mdlData.SelectedScheduleElement.SubjectType != null)
                {
                    //Если количество элементов в списке больше нуля
                    if (cmbSubjectType.Items.Count > 0)
                    {
                        cmbSubjectType.SelectedIndex = mdlData.colSubjectType.IndexOf(mdlData.SelectedScheduleElement.SubjectType);
                    }
                }
                else
                {
                    cmbSubjectType.SelectedIndex = -1;
                }

                //Если специальность указана
                if (mdlData.SelectedScheduleElement.Spec != null)
                {
                    //Если количество элементов в списке больше нуля
                    if (cmbSpecialisation.Items.Count > 0)
                    {
                        cmbSpecialisation.SelectedIndex = mdlData.colSpecialisation.IndexOf(mdlData.SelectedScheduleElement.Spec);
                    }
                }
                else
                {
                    cmbSpecialisation.SelectedIndex = -1;
                }

                //Если номер курса указан
                if (mdlData.SelectedScheduleElement.KursNum != null)
                {
                    //Если количество элементов в списке больше нуля
                    if (cmbKursNum.Items.Count > 0)
                    {
                        cmbKursNum.SelectedIndex = mdlData.colKursNum.IndexOf(mdlData.SelectedScheduleElement.KursNum);
                    }
                }
                else
                {
                    cmbKursNum.SelectedIndex = -1;
                }

                //Аудитория
                txtAuditory.Text = mdlData.SelectedScheduleElement.Auditory;
                //Группа
                txtGroup.Text = mdlData.SelectedScheduleElement.Group;
                //Поток
                txtStream.Text = mdlData.SelectedScheduleElement.Stream;
                //Код редактируемого элемента
                txtCode.Text = mdlData.SelectedScheduleElement.Code.ToString();
            }
        }

        private void FillSubjectList()
        {
            cmbSubject.Items.Clear();
            //Заполняем комбо-список дисциплинами
            for (int i = 0; i <= mdlData.colSubject.Count - 1; i++)
            {
                cmbSubject.Items.Add(mdlData.colSubject[i].Code + ". " + mdlData.colSubject[i].Subject);
            }
        }

        private void FillSubjectTypesList()
        {
            cmbSubjectType.Items.Clear();

            //Заполняем комбо-список видами учебной нагрузки
            for (int i = 0; i <= mdlData.colSubjectType.Count - 1; i++)
            {
                cmbSubjectType.Items.Add(mdlData.colSubjectType[i].Code + ". " + mdlData.colSubjectType[i].Type + " (" + mdlData.colSubjectType[i].Short + ")");
            }
        }

        private void FillSpecialisationList()
        {
            cmbSpecialisation.Items.Clear();

            //Заполняем комбо-список званиями
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                cmbSpecialisation.Items.Add(mdlData.colSpecialisation[i].Code + ". " + mdlData.colSpecialisation[i].ShortUpravlenie + 
                    " (" + mdlData.colSpecialisation[i].ShortInstitute + ")");
            }
        }

        private void FillKursNumList()
        {
            //Очищаем список
            cmbKursNum.Items.Clear();

            //Заполняем комбо-список номерами курсов
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                cmbKursNum.Items.Add(mdlData.colKursNum[i].Kurs);
            }
        }

        //Сохранение сведений
        private void btnSave_Click(object sender, EventArgs e)
        {
            clsSchedule Sch = null;

            if (cmbSubject.SelectedIndex >= 0)
            {
                mdlData.SelectedScheduleElement.Subject = mdlData.colSubject[cmbSubject.SelectedIndex];
            }
            else
            {
                mdlData.SelectedScheduleElement.Subject = null;
            }

            if (cmbSubjectType.SelectedIndex >= 0)
            {
                mdlData.SelectedScheduleElement.SubjectType = mdlData.colSubjectType[cmbSubjectType.SelectedIndex];
            }
            else
            {
                mdlData.SelectedScheduleElement.SubjectType = null;
            }

            if (cmbSpecialisation.SelectedIndex >= 0)
            {
                mdlData.SelectedScheduleElement.Spec = mdlData.colSpecialisation[cmbSpecialisation.SelectedIndex];
            }
            else
            {
                mdlData.SelectedScheduleElement.Spec = null;
            }

            if (cmbKursNum.SelectedIndex >= 0)
            {
                mdlData.SelectedScheduleElement.KursNum = mdlData.colKursNum[cmbKursNum.SelectedIndex];
            }
            else
            {
                mdlData.SelectedScheduleElement.KursNum = null;
            }

            //
            mdlData.SelectedScheduleElement.Auditory = txtAuditory.Text;
            //
            mdlData.SelectedScheduleElement.Group = txtGroup.Text;
            //
            mdlData.SelectedScheduleElement.Stream = txtStream.Text;

            //Если сказано распространить выбранное на обе недели
            if (chkBothWeeks.Checked)
            {
                //Перебираем все элементы расписания
                for (int i = 0; i < mdlData.colSchedule.Count; i++)
                {
                    //Если тот же преподаватель в том же семестре, в тот же день ведёт ту же пару по другой неделе
                    if (mdlData.colSchedule[i].Lecturer.FIO.Equals(mdlData.SelectedScheduleElement.Lecturer.FIO) &
                        mdlData.colSchedule[i].Semestr.SemNum.Equals(mdlData.SelectedScheduleElement.Semestr.SemNum) &
                        mdlData.colSchedule[i].WeekDay.WeekDay.Equals(mdlData.SelectedScheduleElement.WeekDay.WeekDay) &
                        mdlData.colSchedule[i].Time.Time.Equals(mdlData.SelectedScheduleElement.Time.Time) &
                        !mdlData.colSchedule[i].Week.NumberWeek.Equals(mdlData.SelectedScheduleElement.Week.NumberWeek) )
                    {
                        Sch = mdlData.colSchedule[i];
                        break;
                    }
                }

                if (cmbSubject.SelectedIndex >= 0)
                {
                    Sch.Subject = mdlData.colSubject[cmbSubject.SelectedIndex];
                }
                else
                {
                    Sch.Subject = null;
                }

                if (cmbSubjectType.SelectedIndex >= 0)
                {
                    Sch.SubjectType = mdlData.colSubjectType[cmbSubjectType.SelectedIndex];
                }
                else
                {
                    Sch.SubjectType = null;
                }

                if (cmbSpecialisation.SelectedIndex >= 0)
                {
                    Sch.Spec = mdlData.colSpecialisation[cmbSpecialisation.SelectedIndex];
                }
                else
                {
                    Sch.Spec = null;
                }

                if (cmbKursNum.SelectedIndex >= 0)
                {
                    Sch.KursNum = mdlData.colKursNum[cmbKursNum.SelectedIndex];
                }
                else
                {
                    Sch.KursNum = null;
                }

                Sch.Auditory = txtAuditory.Text;
                Sch.Group = txtGroup.Text;
                Sch.Stream = txtStream.Text;

                Sch.Subj = true;
            }

            //Закрыть
            Dispose();
        }

        //Переход к форме обмена нагрузкой
        private void btnSwap_Click(object sender, EventArgs e)
        {
            //Закрыть
            Dispose();
            //
            frmSwapScheduleElement f = new frmSwapScheduleElement();
            //
            f.Owner = Owner;
            f.ShowDialog();
        }
    }
}
