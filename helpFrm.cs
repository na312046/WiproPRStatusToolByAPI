using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PRStatusAPITool
{
    public partial class helpFrm : Form
    {
        public helpFrm()
        {
            InitializeComponent();
        }

        private void btnOk_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lbtnMoreDetails_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string filePath = AppDomain.CurrentDomain.BaseDirectory + "UserManual.pdf";
                if (System.IO.File.Exists(filePath))
                {
                    System.Diagnostics.Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show(" User help manual not available", "CRM :: Pull Request Status API Tool", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                clsStatus.WriteLog("lbtnMoreDetails_LinkClicked()-->" + ex.Message);
            }
        }
    }
}
