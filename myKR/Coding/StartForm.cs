using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using pacEcxelWork;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;
namespace myKR
{
    public partial class StartForm : Form
    {
        public ExcelWork excelWork = null;
        public bool CANCEL = true;

        public StartForm(String box1, String box2)
        {
            InitializeComponent();
            label4.Visible = false;

            textBox1.Text = box1;
            textBox2.Text = box2;
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
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Відкриття поточного навчального плану";
            openFile.Filter = "Excel *.xls|*.xls";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFile.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Відкриття бази даних студентів:";
            openFile.Filter = "Excel *.xls|*.xls";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = openFile.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != null && textBox1.Text != "" && textBox2.Text != null && textBox2.Text != "")
            {

            }
            else
            {
                MessageBox.Show("Усі поля обов'язкові для заповнення!","Попередження");
                return;
            }
            label4.Visible = true;
            if (!System.IO.File.Exists(textBox1.Text))
            {
                MessageBox.Show("заданий вами файл -" + textBox1.Text + " НЕ ІСНУЄ" +
                    " НЕ ІСНУЄ\nВкажіть правильне розташування файлу");
                return;
            }
            if (!System.IO.File.Exists(textBox2.Text))
            {
                MessageBox.Show("заданий вами файл -" + textBox1.Text + 
                    " НЕ ІСНУЄ\nВкажіть правильне розташування файлу");
                return;
            }

            CANCEL = false;
            button3.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            var thread = new Thread(
                () =>
                {
                    AssignLabel("Підключення до поточного плану ...");
                    excelWork = new ExcelWork(textBox1.Text);
                    AssignLabel("Зчитування студентів ...");
                    excelWork.LoadData_StudDB(textBox2.Text);

                    for (int i = 0; i < excelWork.sheetNames_RobPlan.Length; i++)
                    {
                        if (excelWork.sheetNames_RobPlan[i] == null) break;
                        AssignLabel("Занесення предметів групи - " + excelWork.sheetNames_RobPlan[i]);
                        excelWork.LoadData_RobPlan(excelWork.sheetNames_RobPlan[i]);
                    }

                    this.Invoke((MethodInvoker)delegate
                    {
                        // close the form on the forms thread
                        System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                       // bool bl = true;
                        foreach (System.Diagnostics.Process proc in excelProcs)
                        {
                            //if (bl)
                            //    MessageBox.Show("Переконайтеся, що всі застосунки Excel закриті,\n не збережені дані будуть втрачені!", "Уважно!");
                            proc.Kill();
                            //bl = false;
                        }
                        this.Close();
                    });
                    Thread.CurrentThread.Abort();
                });
            thread.Start();
        }


        void AssignLabel(string text)
        {
            if (InvokeRequired)
            {
                BeginInvoke((Action<string>)AssignLabel, text);
                return;
            }
            label4.Text = text;
        }


        public static void tre()
        {
            MessageBox.Show("Done");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlBook = xlApp.Workbooks.Open(@"D:\Книга11.xls");
            Excel.Worksheet xlSheet = xlBook.Worksheets.get_Item(1);

            //xlSheet.Range["A1"].Value = 123;
            xlSheet.Range["A1"].NumberFormatLocal = "##,##";

            xlBook.Save();
            xlBook.Close();
            xlApp.Quit();
        }

        public String[] getTextBox()
        {
            return new String[] { textBox1.Text, textBox2.Text };
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            ExcelWork ex = new ExcelWork(textBox1.Text);
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            ex.ArhiveZvedVid(openFile.FileName);
        }
    }
}
