using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using pacEcxelWork;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR
{
    public partial class MainForm : Form
    {
        public String studDBPath = null;
        public ExcelWork exWork = null;

        public MainForm()
        {
            InitializeComponent();
            label4.Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                            + @"Data\start.txt";
            StreamReader readFile = new StreamReader(path);

            Visible = false;
            StartForm startForm = new StartForm(readFile.ReadLine().Substring(6), readFile.ReadLine().Substring(6));
            readFile.Close();

            startForm.ShowDialog();
            if (startForm.CANCEL) Environment.Exit(-1);

            StreamWriter writeFile = new StreamWriter(path);
            writeFile.WriteLine("[rp] |" + startForm.getTextBox()[0] + "\n"
                              + "[bd] |" + startForm.getTextBox()[1]);
            writeFile.Close();

            studDBPath = startForm.getTextBox()[1];

            exWork = startForm.excelWork;

            for (int i = 0; i < exWork.sheetNames_RobPlan.Length; i++)
            {
                if (exWork.sheetNames_RobPlan[i] == null) break;
                comboBox1.Items.Add(exWork.sheetNames_RobPlan[i]);
            }

            comboBox1.Items.Add("Усі групи");
            comboBox1.Text = comboBox1.Items[0].ToString();
            Visible = true;
        }

        private void reloadSubject()
        {
            comboBox3.Items.Clear();

            int clock = 1;
            if (comboBox2.Text.Equals("2"))
            {
                if (exWork.dsRobPlan.Tables[comboBox1.Text].Columns[5].ColumnName.ToLower().Contains("всього"))
                    clock = 5;
                else clock = 6;
            }

            int kp = clock + 1;
            for (int i = 0; i < exWork.dsRobPlan.Tables[comboBox1.Text].Rows.Count; i++)
            {
                if (!exWork.dsRobPlan.Tables[comboBox1.Text].Rows[i][clock].Equals("0") || !exWork.dsRobPlan.Tables[comboBox1.Text].Rows[i][kp].Equals("0"))
                    comboBox3.Items.Add(exWork.dsRobPlan.Tables[comboBox1.Text].Rows[i][0].ToString());
            }

            comboBox3.Items.Add("Усі предмети");
            comboBox3.Text = comboBox3.Items[0].ToString();
        }

        

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != null)
            {
                for (int i = 0; i < exWork.dsRobPlan.Tables.Count; i++)
                {
                    if (exWork.dsRobPlan.Tables[i].TableName.Equals(comboBox1.Text))
                    {
                        exWork.currentGroupName = comboBox1.Text;
                        reloadSubject();
                    }
                }
            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            reloadSubject();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int groupCount = 1, subjectCount = 1;

            if (comboBox1.Text.Equals("Усі групи")) groupCount = comboBox1.Items.Count - 1;
            if (comboBox3.Text.Equals("Усі предмети")) subjectCount = comboBox3.Items.Count - 1;

            label4.Visible = true;
            label4.Text = "";

            for (int i = 0; i < groupCount; i++)
            {
                if (groupCount > 1)
                {
                    comboBox1.Text = exWork.sheetNames_RobPlan[i];
                    reloadSubject();
                    subjectCount = comboBox3.Items.Count - 1;
                    exWork.currentGroupName = exWork.sheetNames_RobPlan[i];
                }
                else exWork.currentGroupName = comboBox1.Text.ToString();
                
                for (int j = 0; j < subjectCount; j++)
                {
                    if (subjectCount > 1)
                    {
                        comboBox3.Text = comboBox3.Items[j].ToString();
                    }
                    label4.Text = "Створення обліку успішності для групи - " + comboBox1.Text + " з предмету "
                        + comboBox3.Text;
                    exWork.createOblicUspishnosti("123", Convert.ToInt32(comboBox2.Text.ToString()), comboBox3.Text);
                }
            }
            label4.Visible = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int groupCount = 1;
            if (comboBox1.Text.Equals("Усі групи")) groupCount = comboBox1.Items.Count - 1;

            label4.Visible = true;
            label4.Text = "";

            for (int i = 0; i < groupCount; i++)
            {
                if (groupCount > 1)
                {
                    comboBox1.Text = exWork.sheetNames_RobPlan[i];
                    exWork.currentGroupName = exWork.sheetNames_RobPlan[i];
                }
                label4.Text = "Cтворення зведеної відомості за " + comboBox2.Text + " півріччя для групи - " + comboBox1.Text;
                exWork.zvedVidomist(Convert.ToInt32(comboBox2.Text), comboBox1.Text, "");
            }
            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelWork ex = new ExcelWork();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel *.xls|*.xls";
            openFile.Title = "Виберіть зведену відомість за поточне півріччя";
            openFile.FileName = "Зведена відомість успішності за";
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label4.Visible = true;
                label4.Text = "Занесення у архів зведеної відомості";
                ex.ArhiveZvedVid(openFile.FileName);
                label4.Visible = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals("") || comboBox2.Text.Equals("") || comboBox4.Text.Equals(""))
            {
                MessageBox.Show("Заповніть поля 1, 2 і 4");
                return;
            }
            int groupCount = 1;
            if (comboBox1.Text.Equals("Усі групи")) groupCount = comboBox1.Items.Count - 1;

            label4.Visible = true;
            label4.Text = "";

            for (int i = 0; i < groupCount; i++)
            {
                if (groupCount > 1)
                {
                    comboBox1.Text = exWork.sheetNames_RobPlan[i];
                    exWork.currentGroupName = exWork.sheetNames_RobPlan[i];
                }
                label4.Text = "Cтворення зведеної відомості за " + comboBox2.Text + " півріччя для групи - " + comboBox1.Text;
                exWork.zvedVidomist(Convert.ToInt32(comboBox2.Text), comboBox1.Text, comboBox4.Text);
            }
            label4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            bool bl = true;
            foreach (System.Diagnostics.Process proc in excelProcs)
            {
                if (bl)
                    MessageBox.Show("Переконайтеся, що всі застосунки Excel закриті,\n не збережені дані будуть втрачені!", "Уважно!");
                proc.Kill();
                bl = false;
            }
            Environment.Exit(-1);
        }

        private void списокКураторівToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                 + @"Data\Куратори.xls";

            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;
            xlApp.Workbooks.Open(path);
        }
        
    }
}
