using System;
using System.Windows.Forms;

namespace myKR.Coding
{
    public partial class DeleteStudents : Form
    {
        public string StudDbPath;
        public DeleteStudents(string studDbPath)
        {
            StudDbPath = studDbPath;
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
                if (!comboBox1.Text.Equals("ПІ") || !comboBox1.Text.Equals("МТ"))
                {
                    MessageBox.Show("Заповніть коректно текстове поле 1", "Помилка!");
                    return;
                }
                if (comboBox2.Text.Equals("3") && comboBox2.Text.Equals("4")) return;
                MessageBox.Show("Заповніть коректно текстове поле 2", "Помилка!");
                return;
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

        private void deleteStudents_Load(object sender, EventArgs e)
        {

        }
    }
}
