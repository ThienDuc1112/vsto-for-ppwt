using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn1.CustomedComponents
{
    public partial class CustomTextBox : UserControl
    {
        public CustomTextBox()
        {
            InitializeComponent();
        }

        bool isFocused = false;
        private string text = "";
        private bool multipline = false;
        private Color backColor = Color.White;
        private Color foreColor = Color.Black;
        public string TextBox;
        public string customText
        {
            get { return text; }
            set
            {
                text = value;
                this.Invalidate();
            }
        }

        public bool customMultiline
        {
            get { return multipline; }
            set
            {
                multipline = value;
                this.Invalidate();
            }
        }

        private Color customBackColor {
            get { return backColor; }
            set {
                backColor = value;
                this.Invalidate();
            }
        }

        private Color customForeColor
        {
            get { return foreColor; }
            set
            {
                foreColor = value;
                this.Invalidate();
            }
        }

        private async void labelTimer_Tick(object sender, EventArgs e)
        {
            int y = label1.Location.Y;
            if (!isFocused)
            {
                y -= 2;
                label1.Location = new Point(label1.Location.X, y);
                if (y <= 2)
                {
                    isFocused = true;
                    labelTimer.Stop();
                    label1.Font = new Font("Segoi UI", 8);
                    y = 0;
                    label1.ForeColor = Color.Silver;
                }
            }
            else
            {
                y += 2;
                label1.Location = new Point(label1.Location.X, y);

                if (y >= 18)
                {
                    isFocused = true;
                    labelTimer.Stop();
                    label1.Font = new Font("Segoi UI", 10);
                    y = 18;
                    label1.ForeColor = Color.Black;
                }
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                labelTimer.Start();
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                labelTimer.Start();
            }
        }

        private void CustomTextBox_Paint(object sender, PaintEventArgs e)
        {
            label1.Text = customText;
            if (customMultiline) {
                textBox1.Multiline = true;
                textBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                textBox1.Height = this.Height;
            }
            this.backColor = customBackColor;
            this.ForeColor = customForeColor;
            textBox1.BackColor = customBackColor;
            label1.ForeColor = customForeColor;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            TextBox = textBox1.Text;
        }
    }
}
