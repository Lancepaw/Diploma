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
    public partial class frmParams : Form
    {
        public frmParams()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmParams_Load(object sender, EventArgs e)
        {
            //0
            cmbPrintObjects.Items.Add("не выбран");
            //1
            cmbPrintObjects.Items.Add("ведомственная принадлежность");
            //2
            cmbPrintObjects.Items.Add("префикс вуза");
            //3
            cmbPrintObjects.Items.Add("наименование вуза");
            //4
            cmbPrintObjects.Items.Add("суффикс вуза");
            //5
            cmbPrintObjects.Items.Add("кафедра");
            //6
            cmbPrintObjects.Items.Add("оплата ассист.");
            //7
            cmbPrintObjects.Items.Add("оплата ст. преп.");
            //8
            cmbPrintObjects.Items.Add("оплата доцента");
            //9
            cmbPrintObjects.Items.Add("оплата профессора");

            cmbPrintObjects.SelectedIndex = 0;
            txtAverageLoad.Text = mdlData.AverageLoad.ToString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlData.AverageLoad = Convert.ToInt32(txtAverageLoad.Text);

            if (cmbPrintObjects.SelectedIndex > 0)
            {
                switch (cmbPrintObjects.SelectedIndex)
                {
                    //ведомственная принадлежность
                    case 1:
                        mdlData.MinistryName = txtParams.Text;
                        break;
                    //префикс вуза
                    case 2:
                        mdlData.UniversityPrefName = txtParams.Text;
                        break;
                    //наименование вуза
                    case 3:
                        mdlData.UniversityName = txtParams.Text;
                        break;
                    //суффикс вуза
                    case 4:
                        mdlData.UniversitySuffName = txtParams.Text;
                        break;
                    //наименование кафедры
                    case 5:
                        mdlData.DepartmentName = txtParams.Text;
                        break;
                    //
                    case 6:
                        mdlData.PaymentAssist = Convert.ToDouble(txtParams.Text.Replace('.',','));
                        break;
                    //
                    case 7:
                        mdlData.PaymentStPrep = Convert.ToDouble(txtParams.Text.Replace('.', ','));
                        break;
                    //
                    case 8:
                        mdlData.PaymentDocent = Convert.ToDouble(txtParams.Text.Replace('.', ','));
                        break;
                    //
                    case 9:
                        mdlData.PaymentProff = Convert.ToDouble(txtParams.Text.Replace('.', ','));
                        break;
                }
            }

            cmbPrintObjects.SelectedIndex = 0;
        }

        private void cmbPrintObjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPrintObjects.SelectedIndex > 0)
            {
                switch (cmbPrintObjects.SelectedIndex)
                {
                    //наименование ведомства
                    case 1:
                        txtParams.Text = mdlData.MinistryName;
                        break;
                    //префикс вуза
                    case 2:
                        txtParams.Text = mdlData.UniversityPrefName;
                        break;
                    //наименование вуза
                    case 3:
                        txtParams.Text = mdlData.UniversityName;
                        break;
                    //суффикс вуза
                    case 4:
                        txtParams.Text = mdlData.UniversitySuffName;
                        break;
                    //наименование кафедры
                    case 5:
                        txtParams.Text = mdlData.DepartmentName;
                        break;
                    case 6:
                        txtParams.Text = mdlData.PaymentAssist.ToString();
                        break;
                    case 7:
                        txtParams.Text = mdlData.PaymentStPrep.ToString();
                        break;
                    case 8:
                        txtParams.Text = mdlData.PaymentDocent.ToString();
                        break;
                    case 9:
                        txtParams.Text = mdlData.PaymentProff.ToString();
                        break;
                }
            }
            else
            {
                txtParams.Text = "";
            }
        }
    }
}
