using System;
using System.Windows.Forms;

namespace myKR.Coding
{
    public partial class Control : Form
    {
        public static bool IfShow;
        public static int ButtonClick;

        public Control(string text)
        {
            InitializeComponent();
            label1.Text = text;
            IfShow = false;
            ButtonClick = 0;
        }

        public void SetButtonReseachEnabled(bool b)
        {
            button1.Enabled = b;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ButtonClick = 1;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ButtonClick = 2;
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ButtonClick = 3;
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            IfShow = checkBox1.Checked;
        }

        private void Control_Load(object sender, EventArgs e)
        {

        }
    }
}
