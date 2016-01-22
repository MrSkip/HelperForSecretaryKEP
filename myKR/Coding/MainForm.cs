using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public partial class MainForm : Form
    {
        public string StudDbPath;
        public ExcelWork ExWork;

        public MainForm()
        {
            InitializeComponent();
            label4.Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                            + @"Data\start.txt";
            StreamReader readFile = new StreamReader(path);

            Visible = false;
            var readLine = readFile.ReadLine();
            if (readLine != null)
            {
                StartForm startForm = new StartForm(readLine.Substring(6), readLine.Substring(6));
                readFile.Close();

                startForm.ShowDialog();
                if (startForm.Cancel) Environment.Exit(-1);

                StreamWriter writeFile = new StreamWriter(path);
                writeFile.WriteLine("[rp] |" + startForm.GetTextBox()[0] + "\n"
                                    + "[bd] |" + startForm.GetTextBox()[1]);
                writeFile.Close();

                StudDbPath = startForm.GetTextBox()[1];

                ExWork = startForm.ExcelWork;
            }

            foreach (string t in ExWork.SheetNamesRobPlan.TakeWhile(t => t != null))
            {
                comboBox1.Items.Add(t);
            }

            comboBox1.Items.Add("Усі групи");
            comboBox1.Text = comboBox1.Items[0].ToString();
            Visible = true;
        }

        private void ReloadSubject()
        {
            comboBox3.Items.Clear();

            int clock = 1;
            if (comboBox2.Text.Equals("2"))
            {
                clock = ExWork.DsRobPlan.Tables[comboBox1.Text].Columns[5].ColumnName.ToLower().Contains("всього") ? 5 : 6;
            }

            int kp = clock + 1;
            for (int i = 0; i < ExWork.DsRobPlan.Tables[comboBox1.Text].Rows.Count; i++)
            {
                if (!ExWork.DsRobPlan.Tables[comboBox1.Text].Rows[i][clock].Equals("0") || !ExWork.DsRobPlan.Tables[comboBox1.Text].Rows[i][kp].Equals("0"))
                    comboBox3.Items.Add(ExWork.DsRobPlan.Tables[comboBox1.Text].Rows[i][0].ToString());
            }

            comboBox3.Items.Add("Усі предмети");
            comboBox3.Text = comboBox3.Items[0].ToString();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox1.Text)) return;
            for (int i = 0; i < ExWork.DsRobPlan.Tables.Count; i++)
            {
                if (!ExWork.DsRobPlan.Tables[i].TableName.Equals(comboBox1.Text)) continue;
                ExWork.CurrentGroupName = comboBox1.Text;
                ReloadSubject();
            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            ReloadSubject();
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
                    comboBox1.Text = ExWork.SheetNamesRobPlan[i];
                    ReloadSubject();
                    subjectCount = comboBox3.Items.Count - 1;
                    ExWork.CurrentGroupName = ExWork.SheetNamesRobPlan[i];
                }
                else ExWork.CurrentGroupName = comboBox1.Text.ToString();
                
                for (int j = 0; j < subjectCount; j++)
                {
                    if (subjectCount > 1)
                    {
                        comboBox3.Text = comboBox3.Items[j].ToString();
                    }
                    label4.Text = "Створення обліку успішності для групи - " + comboBox1.Text + " з предмету "
                        + comboBox3.Text;
                    ExWork.CreateOblicUspishnosti("123", Convert.ToInt32(comboBox2.Text.ToString()), comboBox3.Text);
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
                    comboBox1.Text = ExWork.SheetNamesRobPlan[i];
                    ExWork.CurrentGroupName = ExWork.SheetNamesRobPlan[i];
                }
                label4.Text = "Cтворення зведеної відомості за " + comboBox2.Text + " півріччя для групи - " + comboBox1.Text;
                ExWork.ZvedVidomist(Convert.ToInt32(comboBox2.Text), comboBox1.Text, "");
            }
            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelWork ex = new ExcelWork();
            OpenFileDialog openFile = new OpenFileDialog
            {
                Filter = "Excel *.xls|*.xls",
                Title = "Виберіть зведену відомість за поточне півріччя",
                FileName = "Зведена відомість успішності за"
            };
            if (openFile.ShowDialog() != DialogResult.OK) return;
            label4.Visible = true;
            label4.Text = "Занесення у архів зведеної відомості";
            ex.ArhiveZvedVid(openFile.FileName);
            label4.Visible = false;
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
                    comboBox1.Text = ExWork.SheetNamesRobPlan[i];
                    ExWork.CurrentGroupName = ExWork.SheetNamesRobPlan[i];
                }
                label4.Text = "Cтворення зведеної відомості за " + comboBox2.Text + " півріччя для групи - " + comboBox1.Text;
                ExWork.ZvedVidomist(Convert.ToInt32(comboBox2.Text), comboBox1.Text, comboBox4.Text);
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

            Excel.Application xlApp = new Excel.Application {Visible = true};
            xlApp.Workbooks.Open(path);
        }
        
    }
}
