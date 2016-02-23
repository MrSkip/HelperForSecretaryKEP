namespace myKR.Coding
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.зведенаВідомістьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.відомістьЗаМісяцьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.занесенняУАрхівToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.облікУспішностіToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вихідToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.додатковоToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.списокКураторівToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.списокСтудентівToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.друкToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button5 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Група";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(40, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Предмет";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Півріччя";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.comboBox1.Location = new System.Drawing.Point(98, 34);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(134, 21);
            this.comboBox1.TabIndex = 3;
            this.comboBox1.TextChanged += new System.EventHandler(this.comboBox1_TextChanged);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "1",
            "2"});
            this.comboBox2.Location = new System.Drawing.Point(98, 61);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(134, 21);
            this.comboBox2.TabIndex = 4;
            this.comboBox2.Text = "1";
            this.comboBox2.TextChanged += new System.EventHandler(this.comboBox2_TextChanged);
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Location = new System.Drawing.Point(98, 88);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(134, 21);
            this.comboBox3.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(125, 151);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(107, 33);
            this.button1.TabIndex = 6;
            this.button1.Text = "Облік успішності";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(250, 27);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(153, 33);
            this.button2.TabIndex = 7;
            this.button2.Text = "Зведена відомість";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 192);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "label4";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(250, 71);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(153, 33);
            this.button4.TabIndex = 10;
            this.button4.Text = "Відомість за місяць";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // comboBox4
            // 
            this.comboBox4.FormattingEnabled = true;
            this.comboBox4.Items.AddRange(new object[] {
            "Січень",
            "Лютий",
            "Березень",
            "Квітень",
            "Травень",
            "Червень",
            "Липень",
            "Серпень",
            "Вересень",
            "Жовтень",
            "Листопад",
            "Грудень"});
            this.comboBox4.Location = new System.Drawing.Point(98, 115);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(134, 21);
            this.comboBox4.TabIndex = 11;
            this.comboBox4.Text = "Вересень";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(42, 118);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Місяць";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.додатковоToolStripMenuItem,
            this.друкToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(411, 24);
            this.menuStrip1.TabIndex = 13;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.зведенаВідомістьToolStripMenuItem,
            this.відомістьЗаМісяцьToolStripMenuItem,
            this.занесенняУАрхівToolStripMenuItem,
            this.облікУспішностіToolStripMenuItem,
            this.вихідToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // зведенаВідомістьToolStripMenuItem
            // 
            this.зведенаВідомістьToolStripMenuItem.Name = "зведенаВідомістьToolStripMenuItem";
            this.зведенаВідомістьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.зведенаВідомістьToolStripMenuItem.Text = "Зведена відомість";
            this.зведенаВідомістьToolStripMenuItem.Click += new System.EventHandler(this.button2_Click);
            // 
            // відомістьЗаМісяцьToolStripMenuItem
            // 
            this.відомістьЗаМісяцьToolStripMenuItem.Name = "відомістьЗаМісяцьToolStripMenuItem";
            this.відомістьЗаМісяцьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.відомістьЗаМісяцьToolStripMenuItem.Text = "Відомість за місяць";
            this.відомістьЗаМісяцьToolStripMenuItem.Click += new System.EventHandler(this.button4_Click);
            // 
            // занесенняУАрхівToolStripMenuItem
            // 
            this.занесенняУАрхівToolStripMenuItem.Name = "занесенняУАрхівToolStripMenuItem";
            this.занесенняУАрхівToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.занесенняУАрхівToolStripMenuItem.Text = "Занесення у архів";
            this.занесенняУАрхівToolStripMenuItem.Click += new System.EventHandler(this.button3_Click);
            // 
            // облікУспішностіToolStripMenuItem
            // 
            this.облікУспішностіToolStripMenuItem.Name = "облікУспішностіToolStripMenuItem";
            this.облікУспішностіToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.облікУспішностіToolStripMenuItem.Text = "Облік успішності";
            this.облікУспішностіToolStripMenuItem.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // вихідToolStripMenuItem
            // 
            this.вихідToolStripMenuItem.Name = "вихідToolStripMenuItem";
            this.вихідToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.вихідToolStripMenuItem.Text = "Вихід";
            this.вихідToolStripMenuItem.Click += new System.EventHandler(this.button5_Click);
            // 
            // додатковоToolStripMenuItem
            // 
            this.додатковоToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.списокКураторівToolStripMenuItem,
            this.списокСтудентівToolStripMenuItem});
            this.додатковоToolStripMenuItem.Name = "додатковоToolStripMenuItem";
            this.додатковоToolStripMenuItem.Size = new System.Drawing.Size(77, 20);
            this.додатковоToolStripMenuItem.Text = "Додатково";
            // 
            // списокКураторівToolStripMenuItem
            // 
            this.списокКураторівToolStripMenuItem.Name = "списокКураторівToolStripMenuItem";
            this.списокКураторівToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.списокКураторівToolStripMenuItem.Text = "Список кураторів";
            this.списокКураторівToolStripMenuItem.Click += new System.EventHandler(this.списокКураторівToolStripMenuItem_Click);
            // 
            // списокСтудентівToolStripMenuItem
            // 
            this.списокСтудентівToolStripMenuItem.Name = "списокСтудентівToolStripMenuItem";
            this.списокСтудентівToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.списокСтудентівToolStripMenuItem.Text = "Видалення студентів";
            // 
            // друкToolStripMenuItem
            // 
            this.друкToolStripMenuItem.Name = "друкToolStripMenuItem";
            this.друкToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            this.друкToolStripMenuItem.Text = "Друк";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(342, 172);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(69, 33);
            this.button5.TabIndex = 14;
            this.button5.Text = "Вихід";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(411, 209);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.comboBox4);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ComboBox comboBox4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem зведенаВідомістьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem відомістьЗаМісяцьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem занесенняУАрхівToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem облікУспішностіToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem вихідToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem додатковоToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem списокКураторівToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem списокСтудентівToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem друкToolStripMenuItem;
        private System.Windows.Forms.Button button5;

    }
}

