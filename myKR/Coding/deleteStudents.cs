using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace myKR.Coding
{
    public partial class deleteStudents : Form
    {
        public String studDBPath = null;
        public deleteStudents(String studDBPath)
        {
            this.studDBPath = studDBPath;
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != null && comboBox2.Text != null)
            {
                if (!comboBox1.Text.ToString().Equals("ПІ") || !comboBox1.Text.ToString().Equals("МТ"))
                {
                    MessageBox.Show("Заповніть коректно текстове поле 1", "Помилка!");
                    return;
                }
                if (!comboBox2.Text.ToString().Equals("3") || !comboBox2.Text.ToString().Equals("4"))
                {
                    MessageBox.Show("Заповніть коректно текстове поле 2", "Помилка!");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Усі поля обов'язкові для заповнення", "Помилка!");
                return;
            }
            //
            //як саме записуватиметься курс у бд - арабськими чи римськими символами
            ///
            //
        }
    }
}
