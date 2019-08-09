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
    public partial class frmGAKCalc : Form
    {
        public frmGAKCalc()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCount_Click(object sender, EventArgs e)
        {
            //Захватываем общее количество часов на дипломное
            //проектирование из соответствующего текст-бокса
            int General;
            //Захватываем количество студентов-дипломников
            //из соответствующего текст-бокса
            int Students;

            double OneStudentLoad;

            int Main;

            int Outer;

            int Inner;

            int Secretary;

            int Subscribe;
            int ToMain;
            int ToOuter;
            int ToInner;
            int ToSecretary;
            int People;
            int Diploma;
            int GAK;

            if (!(txtGeneral.Text == ""))
            {
                General = Convert.ToInt32(txtGeneral.Text);
            }
            else
            {
                General = 0;
            }

            if (!(txtStudNum.Text == ""))
            {
                Students = Convert.ToInt32(txtStudNum.Text);
            }
            else
            {
                Students = 0;
            }

            if (!(txtMain.Text == ""))
            {
                Main = Convert.ToInt32(txtMain.Text);
            }
            else
            {
                Main = 0;
            }

            if (!(txtOuter.Text == ""))
            {
                Outer = Convert.ToInt32(txtOuter.Text);
            }
            else
            {
                Outer = 0;
            }

            if (!(txtInner.Text == ""))
            {
                Inner = Convert.ToInt32(txtInner.Text);
            }
            else
            {
                Inner = 0;
            }

            if (!(txtSecr.Text == ""))
            {
                Secretary = Convert.ToInt32(txtSecr.Text);
            }
            else
            {
                Secretary = 0;
            }

            Subscribe = Students;
            txtSubscribe.Text = Subscribe.ToString();

            ToMain = Students * Main;
            txtToMain.Text = ToMain.ToString();

            ToOuter = Outer * Students;
            txtToOuter.Text = ToOuter.ToString();

            ToOuter = Outer * Students;
            txtToOuter.Text = ToOuter.ToString();

            ToInner = Convert.ToInt32(Inner * Students * 0.5);
            txtToInner.Text = ToInner.ToString();

            ToSecretary = Convert.ToInt32(Secretary * Students * 0.5);
            txtToSecretary.Text = ToSecretary.ToString();

            People = Inner + Outer + Main + Secretary;
            txtPeople.Text = People.ToString();

            GAK = Convert.ToInt32(ToOuter + ToInner + ToSecretary + ToMain + Subscribe);
            txtGAK.Text = GAK.ToString();

            Diploma = General - GAK;
            txtDiploma.Text = Diploma.ToString();

            OneStudentLoad = Diploma / Students;
            txtStudentLoad.Text = OneStudentLoad.ToString();
            
            txtLect.Text = ToSecretary.ToString();      
        }

        private void frmGAKCalc_Load(object sender, EventArgs e)
        {
            txtSubscribe.Enabled = false;
            txtToMain.Enabled = false;
            txtToInner.Enabled = false;
            txtToOuter.Enabled = false;
            txtToSecretary.Enabled = false;
            txtStudentLoad.Enabled = false;

            txtDiploma.Enabled = false;
            txtGAK.Enabled = false;

            txtPeople.Enabled = false;
            txtLect.Enabled = false;   

            FillDiplomaList(mdlData.colDistribution);
        }

        private void FillDiplomaList(IList<clsDistribution> collection)
        {
            //Очистка комбо-списка распределения нагрузки
            cmbDiplomaList.Items.Clear();

            if (!(collection == null))
            {
                //Заполняем комбо-список распределения нагрузки
                for (int i = 0; i <= collection.Count - 1; i++)
                {
                    if (collection[i].DiplomaPaper > 0 & collection[i].EnteredHours > 0)
                    {
                        if (!(collection[i].Speciality == null))
                        {
                            cmbDiplomaList.Items.Add(collection[i].Code + ". "
                                                     + collection[i].Speciality.ShortDop);
                        }
                    }
                }

                if (cmbDiplomaList.Items.Count > 0)
                {
                    cmbDiplomaList.SelectedIndex = 0;
                }
                else
                {
                    cmbDiplomaList.SelectedIndex = -1;
                }
            }
        }

        private void cmbDiplomaList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Spec;
            int Students;
            int EnteredHours;

            Spec = "";
            Students = 0;
            EnteredHours = 0;

            //Перебираем все строки таблицы распределения нагрузки
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //Фиксируем название специальности текущего элемента
                if (!(mdlData.colDistribution[i].Speciality == null))
                {
                    Spec = mdlData.colDistribution[i].Speciality.ShortUpravlenie;
                }
                else
                {
                    Spec = "";
                }

                //Если название специальности выбранного из списка элемента совпадает
                //c названием специальности текущего элемента
                if (
                    (cmbDiplomaList.SelectedItem.ToString().EndsWith(Spec)) &
                    (cmbDiplomaList.SelectedItem.ToString().Substring
                          (
                            cmbDiplomaList.SelectedItem.ToString().IndexOf(".") + 2
                          )
                          .Length == Spec.Length)
                   )
                {
                    //Если есть часы на дипломное проектирование
                    //то увеличиваем счётчик студентов
                    if (mdlData.colDistribution[i].DiplomaPaper > 0)
                    {
                        if (!(mdlData.colDistribution[i].Subject.Subject.Contains("Подпись")))
                        {
                            Students += 1;
                        }
                    }

                    //Если указаны общие часы на ГАК или Дипломное проектирование
                    //Суммируем эти часы в единый набор

                    if ((mdlData.colDistribution[i].DiplomaPaper > 0) || (mdlData.colDistribution[i].GAK > 0))
                    {
                        EnteredHours += mdlData.colDistribution[i].EnteredHours;
                    }
                }
            }

            txtGeneral.Text = EnteredHours.ToString();
            txtStudNum.Text = Students.ToString();
        }
    }
}
