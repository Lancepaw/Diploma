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
    public partial class frmLectInput : Form
    {
        public frmLectInput()
        {
            InitializeComponent();
        }

        private void frmLectInput_Load(object sender, EventArgs e)
        {
            FillLecturerList();
        }

        private void FillLecturerList()
        {
            //Очищаем список
            cmbLecturer.Items.Clear();

            //Заполняем комбо-список преподавателям
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturer.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }
        }

        private void btnEnter_Click(object sender, EventArgs e)
        {
            mdlData.SelectedLecturer = mdlData.colLecturer[cmbLecturer.SelectedIndex];
            this.Close();
        }
    }
}
