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
    public partial class frmDistributionAccept : Form
    {
        public frmDistributionAccept()
        {
            InitializeComponent();
        }

        private void frmDistributionAccept_Load(object sender, EventArgs e)
        {
            clsDistributionDetailed DD;

            cmbInput.Items.Clear();
            for (int i = 0; i < mdlData.colDistributionDetailed.Count; i++)
            {
                DD = mdlData.colDistributionDetailed[i];

                if (DD.Lecturer != null)
                {
                    if (mdlData.SelectedLecturer.FIO.Equals(DD.Lecturer.FIO))
                    {
                        if (DD.SubjHours > 0)
                        {
                            cmbInput.Items.Add(i + ". " + DD.Speciality.ShortInstitute + "-" + DD.KursNum.Kurs
                                + " (" + DD.SubjType.Short + ") - " + DD.Subject.Subject);
                        }
                    }
                }
            }
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            mdlData.SelectedScheduleElement.Link =
                mdlData.colDistributionDetailed[Convert.ToInt32(cmbInput.SelectedItem.ToString().Substring(
                    0, cmbInput.SelectedItem.ToString().IndexOf('.')))];
            mdlData.SelectedScheduleElement.Auditory = txtAuditory.Text;
            this.Close();
        }
    }
}
