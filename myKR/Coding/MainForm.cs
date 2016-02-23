using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            label4.Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Visible = false;
            string path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                            + @"Data\start.txt";
            StreamReader readFile = new StreamReader(path);

            StartForm startForm = new StartForm(readFile.ReadLine().Substring(13), readFile.ReadLine().Substring(13));

            readFile.Close();

            startForm.ShowDialog();

            if (startForm.Cancel) Environment.Exit(-1);

            StreamWriter writeFile = new StreamWriter(path);
            writeFile.WriteLine("[work plan] |" + startForm.GetTextBox()[0] + "\n"
                                + "[data base] |" + startForm.GetTextBox()[1]);
            writeFile.Close();

            foreach (Group @group in Manager.Groups)
            {
                comboBox1.Items.Add(group.Name);
            }

            comboBox1.Items.Add("Усі групи");
            comboBox1.Text = comboBox1.Items[0].ToString();
            Visible = true;
        }

        private void ReloadSubject()
        {
            if (comboBox1.Items.Count == 0) return;

            comboBox3.Items.Clear();

            if (comboBox1.Text.Equals(comboBox1.Items[comboBox1.Items.Count - 1]))
                return;

            foreach (Subject subject in Manager.GetGroupByName(comboBox1.Text).Subjects)
            {
                Semestr semestr =
                    comboBox2.Text.Equals("1") ? subject.FirstSemestr : subject.SecondSemestr;
                if (semestr != null)
                    comboBox3.Items.Add(subject.Name);
            }

            comboBox3.Items.Add("Усі предмети");
            comboBox3.Text = comboBox3.Items[0].ToString();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            ReloadSubject();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            ReloadSubject();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(comboBox1.Text))
                return;

            label4.Visible = true;
            label4.Text = comboBox1.Text.Equals("Усі групи") || comboBox3.Text.Equals("Усі предмети")
                ? "Працюю, створення обліків успішності ..." : "Працюю, створення обліку успішності ...";

            Manager.CreateOblicUspishnosti(comboBox1.Text.Equals("Усі групи") ? "" : comboBox1.Text,
                comboBox3.Text.Equals("Усі предмети") ? "" : comboBox3.Text, Int32.Parse(comboBox2.Text));

            label4.Visible = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox1.Text) || string.IsNullOrEmpty(comboBox3.Text)) 
                return;

            label4.Visible = true;
            label4.Text = comboBox1.Text == "Усі групи"
                ? "Працюю, створення зведених відомостей успішності ..."
                : "Працюю, створення зведеної відомості успішності ...";

            label4.Text = comboBox1.Text.Equals("Усі групи")
                ? "Працюю, створення зведених відомостей ..." : "Працюю, створення зведеної відомості ...";

            Manager.CreateVidomistUspishnosti(comboBox1.Text, Int32.Parse(comboBox2.Text));

            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
//            ExcelWork ex = new ExcelWork();
//            OpenFileDialog openFile = new OpenFileDialog
//            {
//                Filter = "Excel *.xls|*.xls",
//                Title = "Виберіть зведену відомість за поточне півріччя",
//                FileName = "Зведена відомість успішності за"
//            };
//            if (openFile.ShowDialog() != DialogResult.OK) return;
//            label4.Visible = true;
//            label4.Text = "Занесення у архів зведеної відомості";
//            ex.ArhiveZvedVid(openFile.FileName);
//            label4.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
//            if (comboBox1.Text.Equals("") || comboBox2.Text.Equals("") || comboBox4.Text.Equals(""))
//            {
//                MessageBox.Show("Заповніть поля 1, 2 і 4");
//                return;
//            }
//            int groupCount = 1;
//            if (comboBox1.Text.Equals("Усі групи")) groupCount = comboBox1.Items.Count - 1;
//
//            label4.Visible = true;
//            label4.Text = "";
//
//            for (int i = 0; i < groupCount; i++)
//            {
//                if (groupCount > 1)
//                {
//                    comboBox1.Text = ExWork.SheetNamesRobPlan[i];
//                    ExWork.CurrentGroupName = ExWork.SheetNamesRobPlan[i];
//                }
//                label4.Text = "Cтворення зведеної відомості за " + comboBox2.Text + " півріччя для групи - " + comboBox1.Text;
//                ExWork.ZvedVidomist(Convert.ToInt32(comboBox2.Text), comboBox1.Text, comboBox4.Text);
//            }
            label4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
//            ExcelFile.App.Quit();
//            System.Diagnostics.Process[] excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
//            bool bl = true;
//            foreach (System.Diagnostics.Process proc in excelProcs)
//            {
//                if (bl)
//                    MessageBox.Show("Переконайтеся, що всі застосунки Excel закриті,\n не збережені дані будуть втрачені!", "Уважно!");
//                proc.Kill();
//                bl = false;
//            }
            Environment.Exit(-1);
        }

        private void списокКураторівToolStripMenuItem_Click(object sender, EventArgs e)
        {
//            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
//                                 + @"Data\Куратори.xls";
//
//            Excel.Application xlApp = new Excel.Application {Visible = true};
//            xlApp.Workbooks.Open(path);
        }
        
    }
}
