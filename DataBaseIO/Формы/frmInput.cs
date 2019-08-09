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
    public partial class frmInput : Form
    {
        public frmInput()
        {
            InitializeComponent();
        }

        private void txtRate_TextChanged(object sender, EventArgs e)
        {

        }

        private void frmInput_Load(object sender, EventArgs e)
        {
            //cmbDuty.Items.Add("не указана");
            //for (int i = 0; i <= mdlData.colDuty.Count - 1; i++)
            //{
            //    cmbDuty.Items.Add(mdlData.colDuty[i].Name);
            //}
            //cmbDuty.SelectedIndex = 0;
            
            //if (mdlData.flgInput == 1)
            //{
            //    btnCommand.Text = "Изменить";
            //    txtSurname.Text = mdlData.colWorker[mdlData.currentCode].Surname;                
            //    txtName.Text = mdlData.colWorker[mdlData.currentCode].Name;                
            //    txtPatronymic.Text = mdlData.colWorker[mdlData.currentCode].Patronymic;                
            //    txtPassport.Text = mdlData.colWorker[mdlData.currentCode].Passport;                
            //    cmbDuty.SelectedIndex = mdlData.colWorker[mdlData.currentCode].Duty;
            //    txtRate.Text = Convert.ToString(mdlData.colWorker[mdlData.currentCode].Rate);
            //}

            //if (mdlData.flgInput == 2)
            //{
            //    btnCommand.Text = "Добавить";

            //}
        }

        private void btnCommand_Click(object sender, EventArgs e)
        {
            
            //if (mdlData.flgInput == 2)
            //{              
            //    mdlData.colWorker.Add(new clsRaspred());
            //    mdlData.currentCode = mdlData.colWorker.Count - 1;
            //    mdlData.colWorker[mdlData.currentCode].Code = mdlData.colWorker.Count;
            //}

            //mdlData.colWorker[mdlData.currentCode].Surname = txtSurname.Text;
            //mdlData.colWorker[mdlData.currentCode].Name = txtName.Text;
            //mdlData.colWorker[mdlData.currentCode].Patronymic = txtPatronymic.Text;
            //mdlData.colWorker[mdlData.currentCode].Passport = txtPassport.Text;
            //mdlData.colWorker[mdlData.currentCode].Duty = cmbDuty.SelectedIndex;
            //mdlData.colWorker[mdlData.currentCode].Rate = Convert.ToDouble(txtRate.Text);
            this.Close();
        }
    }
}