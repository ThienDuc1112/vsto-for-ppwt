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
    public partial class noteForm : Form
    {
        public string Note { get; set; }
        public noteForm()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Note = tbNote.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }
}
