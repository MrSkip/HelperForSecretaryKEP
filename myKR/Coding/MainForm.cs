using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace myKR.Coding
{
    public partial class MainForm : Form
    {
        private PathsFile PathsFile = PathsFile.GetPathsFile();

        public MainForm()
        {
            InitializeComponent();
            label4.Visible = false;
            Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Visible = false;

            StartForm startForm = new StartForm();
            startForm.ShowDialog();

            if (startForm.Cancel)
            {
                Environment.Exit(-1);
            }

            PathsFile.PathsDto.PathToWorkPlan = startForm.GetTextBox()[0];
            PathsFile.PathsDto.PathToStudentDb = startForm.GetTextBox()[1];

            PathsFile.WriteFromObjectToJson();

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
            if (!comboBox2.Text.Equals("1") && !comboBox2.Text.Equals("2")) return;

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

            Manager.CreateVidomistUspishnosti(comboBox1.Text, Int32.Parse(comboBox2.Text), null);

            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox1.Text) || string.IsNullOrEmpty(comboBox2.Text)
                || string.IsNullOrEmpty(comboBox4.Text))
            {
                MessageBox.Show("Заповніть поля 1, 2 і 4");
                return;
            }

            label4.Visible = true;
            label4.Text = comboBox1.Text.Equals("Усі групи")
                ? "Працюю, створення зведених відомостей за місяць" : "Працюю, створення зведеної відомості за місяць";

            Manager.CreateVidomistUspishnosti(comboBox1.Text, Int32.Parse(comboBox2.Text), comboBox4.Text);

            label4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Manager.CloseMainExcelApp();
            Environment.Exit(-1);
        }

        private void списокКураторівToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox1.Text))
                return;

            label4.Visible = true;
            label4.Text = comboBox1.Text.Equals("Усі групи")
                ? "Працюю, створення атесатів" : "Працюю, створення атесату";

            List<string> list = new List<string>();
            if (comboBox1.Text.Equals("Усі групи"))
            {
                for (byte i = 0; i < comboBox1.Items.Count - 1; i++)
                {
                    list.Add(comboBox1.Items[i].ToString());
                }
            }
            else
            {
                list.Add(comboBox1.Text);
            }

            Manager.CreateAtestat(list);

            label4.Visible = false;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Manager.CloseMainExcelApp();
            ExcelApplication.ExcelApplication.Kill(ExcelApplication.ExcelApplication.App);
        }
        
    }
}
