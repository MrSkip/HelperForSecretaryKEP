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

            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
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

            comboBox1.SelectedIndex = 0;

            comboBox1.Items.Add("Усі групи");
            Visible = true;
        }

        private void ReloadSubject()
        {
            if (comboBox1.Items.Count == 0) return;

            if (comboBox1.SelectedItem.Equals(comboBox1.Items[comboBox1.Items.Count - 1]))
            {
                comboBox3.SelectedIndex = comboBox3.Items.Count - 1;
                return;
            }

            comboBox3.Items.Clear();

            foreach (Subject subject in Manager.GetGroupByName(comboBox1.SelectedItem.ToString()).Subjects)
            {
                Semestr semestr =
                    comboBox2.SelectedItem.ToString().Equals("1") ? subject.FirstSemestr : subject.SecondSemestr;
                if (semestr != null)
                    comboBox3.Items.Add(subject.Name);
            }
            
            comboBox3.Items.Add("Усі предмети");
            comboBox3.SelectedIndex = 0;
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
            if (comboBox1.Items.Count == 0)
                return;

            label4.Visible = true;
            label4.Text = comboBox1.SelectedItem.ToString().Equals("Усі групи")
                || comboBox3.SelectedItem.ToString().Equals("Усі предмети")
                ? "Працюю, створення обліків успішності ..." : "Працюю, створення обліку успішності ...";

            Manager.CreateOblicUspishnosti(comboBox1.SelectedItem.ToString().Equals("Усі групи") 
                    ? "" 
                    : comboBox1.SelectedItem.ToString(),
                comboBox3.SelectedItem.ToString().Equals("Усі предмети") 
                    ? ""
                    : comboBox3.SelectedItem.ToString(), Int32.Parse(comboBox2.SelectedItem.ToString()));

            label4.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count == 0 || comboBox3.Items.Count == 0)
                return;

            label4.Visible = true;
            label4.Text = comboBox1.SelectedItem.ToString() == "Усі групи"
                ? "Працюю, створення зведених відомостей успішності ..."
                : "Працюю, створення зведеної відомості успішності ...";

            label4.Text = comboBox1.SelectedItem.ToString().Equals("Усі групи")
                ? "Працюю, створення зведених відомостей ..." : "Працюю, створення зведеної відомості ...";

            Manager.CreateVidomistUspishnosti(comboBox1.SelectedItem.ToString(), Int32.Parse(comboBox2.SelectedItem.ToString()), null);

            label4.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count == 0 || comboBox2.Items.Count == 0)
            {
                MessageBox.Show("Заповніть поля 1, 2");
                return;
            }

            label4.Visible = true;
            label4.Text = comboBox1.SelectedItem.ToString().Equals("Усі групи")
                ? "Працюю, створення зведених відомостей за місяць" : "Працюю, створення зведеної відомості за місяць";

            Manager.CreateVidomistUspishnosti(comboBox1.SelectedItem.ToString(), Int32.Parse(comboBox2.SelectedItem.ToString()),
                comboBox4.SelectedItem.ToString());

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
            if (comboBox1.Items.Count == 0)
                return;

            label4.Visible = true;
            label4.Text = comboBox1.SelectedItem.ToString().Equals("Усі групи")
                ? "Працюю, створення атестатів" : "Працюю, створення атестату";

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
                list.Add(comboBox1.SelectedItem.ToString());
            }

            Manager.CreateAtestat(list);

            label4.Visible = false;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Manager.CloseMainExcelApp();
            ExcelApplication.ExcelApplication.Kill(ExcelApplication.ExcelApplication.App);
        }

        private void SelectedIndexChanged(object sender, EventArgs e)
        {
            ReloadSubject();
        }
        
    }
}
