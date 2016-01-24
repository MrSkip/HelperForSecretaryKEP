using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace myKR.Coding
{
    public partial class StartForm : Form
    {
        public ExcelWork ExcelWork;
        public bool Cancel = true;

        public StartForm(string field1, string fiedl2)
        {
            InitializeComponent();
            label4.Visible = false;

            textBox1.Text = field1;
            textBox2.Text = fiedl2;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process proc in excelProcs)
            {
                proc.Kill();
            }
            Environment.Exit(-1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Title = "Відкриття поточного навчального плану",
                Filter = "Excel *.xls|*.xls"
            };
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFile.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Title = "Відкриття бази даних студентів:",
                Filter = "Excel *.xls|*.xls"
            };
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFile.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Усі поля обов'язкові для заповнення!", "Попередження");
                return;
            }
            if (!File.Exists(textBox1.Text))
            {
                MessageBox.Show("заданий вами файл -" + textBox1.Text + " НЕ ІСНУЄ" +
                    "\nВкажіть правильне розташування файлу");
                return;
            }
            if (!File.Exists(textBox2.Text))
            {
                MessageBox.Show("заданий вами файл -" + textBox1.Text + 
                    " НЕ ІСНУЄ\nВкажіть правильне розташування файлу");
                return;
            }

            label4.Visible = true;
            Cancel = false;
            button3.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            
            var thread = new Thread(
                () =>
                {
                    AssignLabel("Зачекайте, будь ласка, зчитуються дані ...");
                    Manager.ReadData(textBox1.Text, textBox2.Text);
                    Invoke((MethodInvoker)Close);
                    Thread.CurrentThread.Abort();
                });
            thread.Start();
        }


        private void AssignLabel(string text)
        {
            if (InvokeRequired)
            {
                BeginInvoke((Action<string>)AssignLabel, text);
                return;
            }
            label4.Text = text;
        }

        public string[] GetTextBox()
        {
            return new[] { textBox1.Text, textBox2.Text };
        }

        private void StartForm_Load(object sender, EventArgs e)
        {

        }
    }
}
