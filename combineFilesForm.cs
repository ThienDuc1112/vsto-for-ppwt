using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class combineFilesForm : Form
    {
        public string filePath1 { get; set; }
        public string filePath2 { get; set; }
        public combineFilesForm()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            filePath1 = SelectPowerPointFile();
            if (!string.IsNullOrEmpty(filePath1))
            {
                textBox1.Text = filePath1;
            }
            else
            {
                MessageBox.Show("Error while selecting files");
            }
        }

        private void btnImport2_Click(object sender, EventArgs e)
        {
            filePath2 = SelectPowerPointFile();
            if (!string.IsNullOrEmpty(filePath2))
            {
                textBox2.Text = filePath2;
            }
            else
            {
                MessageBox.Show("Error while selecting files");
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private string SelectPowerPointFile()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PowerPoint files (*.pptx)|*.pptx";
            ofd.Title = "Select a PowerPoint file";

            if(ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }

            return null;
        }
    }
}
