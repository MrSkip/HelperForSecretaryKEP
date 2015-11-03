using System;
using System.Data;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;


namespace pacEcxelWork
{
    public class ExcelWork
    {
        private Excel.Application xlApp = null;
        private Excel.Workbook xlWorkBook = null;
        private Excel.Worksheet xlWorkSheet = null;
        private object misValue = System.Reflection.Missing.Value;

        //Path to Excel file
        private String pathRobPlan = null;

        // where argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString() - "Напрям підготовки", .[1] - "Спеціальність", .[2] - "Курс", .[3] - "рік"
        // .[4] - перше півріччя, .[5] - друге півріччя, .[6] - код спеціальності
        public DataSet argDataSet = new DataSet();
        public String[] robPlanArgs = new String[7];

        //Group names of sheets from excel file "PlanRob"
        public String[] sheetNames_RobPlan = new String[20];

        //in this DataSet we load data from robBlan and StudDB
        public DataSet dsRobPlan = new DataSet();

        //
        public String currentGroupName = null;

        public ExcelWork(String pathRobPlan)
        {
            this.pathRobPlan = pathRobPlan;
            LoadSheetName_RobPlan();
        }
        public ExcelWork()
        {

        }

        private void LoadSheetName_RobPlan()
        {
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(pathRobPlan);

            int countGroup = 0;
            for (int i = 1; i <= xlWorkBook.Sheets.Count; i++)
            {
                String name = xlWorkBook.Worksheets.get_Item(i).Name;
                if (name.Length == 8 && name.IndexOf('-') == 2)
                {
                    sheetNames_RobPlan[countGroup] = xlWorkBook.Worksheets.get_Item(i).Name;
                    countGroup++;
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }

        //Read need argument for our program from Ecxel file
        public void LoadData_RobPlan(String currentGroupName)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(pathRobPlan);

            this.currentGroupName = currentGroupName;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(currentGroupName);

            // where argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString() - "Напрям підготовки", .[1] - "Спеціальність", .[2] - "Курс", .[3] - "рік"
            // .[4] - перше півріччя, .[5] - друге півріччя, .[6] - код спеціальності
            try
            {
                argDataSet.Tables.Add(currentGroupName);
            }
            catch (Exception ex)
            {
                xlWorkBook.Close();
                xlApp.Quit();
                return;
            }

            argDataSet.Tables[currentGroupName].Columns.Add("Напрям підготовки");
            argDataSet.Tables[currentGroupName].Columns.Add("Спеціальність");
            argDataSet.Tables[currentGroupName].Columns.Add("Курс");
            argDataSet.Tables[currentGroupName].Columns.Add("Рік");
            argDataSet.Tables[currentGroupName].Columns.Add("Перше півріччя");
            argDataSet.Tables[currentGroupName].Columns.Add("Друге півріччя");
            argDataSet.Tables[currentGroupName].Columns.Add("Код спеціальності");

            DataRow argNewRow = argDataSet.Tables[currentGroupName].NewRow();
            argDataSet.Tables[currentGroupName].Rows.Add(argNewRow);

            //Read "Напряму підготовки"
            String arg = xlWorkSheet.Cells[6, "R"].Value.ToString().Trim();
            argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"] = arg.Substring(arg.IndexOf("\"") + 1, arg.LastIndexOf("\"") - arg.IndexOf("\"") - 1);

            //Read "Спеціальність"
            arg = xlWorkSheet.Cells[7, "R"].Value;
            argDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"] = arg.Substring(arg.IndexOf("\"") + 1, arg.LastIndexOf("\"") - arg.IndexOf("\"") - 1);
            
            //Код спеціальності
            arg = arg.Trim().Substring(0, arg.Trim().IndexOf(" \""));
            argDataSet.Tables[currentGroupName].Rows[0]["Код спеціальності"] = arg.Substring(arg.IndexOf(" "));

            //read "Курс"
            arg = xlWorkSheet.Cells[9 , "R"].Value.Trim();
            arg = arg.Substring(7);
            argDataSet.Tables[currentGroupName].Rows[0]["Курс"] = arg.Remove(arg.IndexOf("_"));

            //Read "Рік"
            argDataSet.Tables[currentGroupName].Rows[0]["Рік"] = xlWorkSheet.Cells[6, "B"].Value.ToString().Substring(xlWorkSheet.Cells[6, "B"].Value.ToString().Length - 9, 4);

            
            //Read need values from `RobPlan` to our `dsTable`
            int xlIterator = 14;
            int countFormControl = 2;

            String[] zaput = new String[12];
            int iZaput = 0;

            DataTable countryPas = new DataTable();
            countryPas.Columns.Add("pas");
            countryPas.Columns.Add("term");

            while (true)
            {
                xlIterator++;

                String value = xlWorkSheet.Cells[xlIterator, "C"].Value;
                //MessageBox.Show(value);
                if (value == null) continue;
                else if (!value.Equals("Назви навчальних  дисциплін") && xlIterator == 15)
                {
                    MessageBox.Show("Назви значень полів у:\n" + pathRobPlan + "\nне співпадає із заданами значеннями у програмі\n" +
                         "Джерело - " + currentGroupName + "\n У клітині \"С15\" очікувалося значення 'Назви навчальних  дисциплін'");
                    xlWorkBook.Close();
                    xlApp.Quit();
                    return;
                }
                else if (value.Trim().Equals("Разом")) break;

                if (xlIterator == 15)
                {
                    //Задаємо усі можиливі місця розміщення клітини з назвою про державний екзамен  - "Назва"
                    //[][0] - колонка з назвою екзамена, [][1] - колонка з рядком езамену, [][2] - колонка з назвою семестра проходження екзамену
                    //відповідно, назва екзамену і назва семестру будуть знаходитися у одному рядку
                    String[][] doubleCell = {
                                                new String[] {"BE", "49", "BO"},
                                                new String[] {"BE", "47", "BO"},
                                                new String[] {"BE", "42", "BO"},
                                                new String[] {"BE", "43", "BO"},
                                                new String[] {"AX", "39", "BC"}
                                            };
                    //Цикл для проходження усіх можливих місць знаходження назви державного екзамену, клітини із значенням "Назва"
                    for (int i = 0; i < doubleCell.Length; i++)
                    {
                        //Для унеможливлення винекнинне помилок автоматично присвоюємо тип змінній howType
                        //Якщо у нашій клітині є якесь значення, то виконується умова
                        var howType = xlWorkSheet.Cells[doubleCell[i][1], doubleCell[i][0]].Value;
                        if (howType != null)
                        {
                            //Перевіряємо чи це справді клітина з назвою про державний екзамен
                            if (howType.ToString().Trim().Equals("Назва"))
                            {
                                //Назви полів залишатимуться сталимим, натомість рядок постійно інкрементуватиметься, тому переводиму його до типу int
                                int iter = System.Convert.ToInt32(doubleCell[i][1]) + 1;
                                //Коли знайдено місце знаходження потрібної нам клітини, ми шукаємо чи є хоча б якісь екзамени
                                while (true)
                                {
                                    var whatType = xlWorkSheet.Cells[iter, doubleCell[i][0]].Value;
                                    if (whatType != null)
                                    {
                                        //Якщо екзамен знайдено то записуємо у таблицю його назву і семестр у якому проводитиметься екзамен
                                        countryPas.Rows.Add(whatType,
                                            arabNormalize(xlWorkSheet.Cells[iter, doubleCell[i][2]].Value.ToString().Trim()));
                                        iter++;
                                    }
                                    else break;
                                }
                                break;
                            }
                        }
                    }
                    //Зчитування практики, якщо практика існує для поточного семестру то заносимо її у масив Mas
                    //Create table "currentTable"
                    dsRobPlan.Tables.Add(currentGroupName);

                    //Add to table column "Назви навчальних  дисциплін"
                    dsRobPlan.Tables[currentGroupName].Columns.Add(value);
                    zaput[iZaput] = "C";
                    iZaput++;

                    

                    //непарний семестр
                    argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"] = xlWorkSheet.Cells[xlIterator, "Y"].Value.ToString().Substring(0,
                        xlWorkSheet.Cells[xlIterator, "Y"].Value.ToString().IndexOf(" "));
                    argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"] = arabNormalize(argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"].ToString());

                    //Create string name of current term
                    String term = " [семестр " + argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"].ToString() + "]";

                    //Add to table column "всього годин"
                    dsRobPlan.Tables[currentGroupName].Columns.Add(xlWorkSheet.Cells[16, "Y"].Value + " годин" + term);
                    zaput[iZaput] = "Y";
                    iZaput++;

                    //Add to table column "курсові роботи, проекти"
                    dsRobPlan.Tables[currentGroupName].Columns.Add("КП" + term);
                    zaput[iZaput] = "AK";
                    iZaput++;

                    //Add to table column "екзамен"
                    dsRobPlan.Tables[currentGroupName].Columns.Add("Іспит" + term);
                    zaput[iZaput] = "AO";
                    iZaput++;

                    //Add to table column "диф  залік" or "залік"
                    String s = xlWorkSheet.Cells[18, "AQ"].Value.ToString();
                    String s2 = null;
                    if (s.Trim().Equals("диф  залік")) s2 = "Д/З";
                    else s2 = "Залік";
                    dsRobPlan.Tables[currentGroupName].Columns.Add(s2 + term);
                    zaput[iZaput] = "AQ";
                    iZaput++;

                    //Add to table column "диф  залік" - if exist
                    if (xlWorkSheet.Cells[18, "AR"].Value != null) 
                    {
                        dsRobPlan.Tables[currentGroupName].Columns.Add("Д/З" + term);
                        countFormControl = 3;
                        zaput[iZaput] = "AR";
                        iZaput++;
                    }

                    //парний семестр
                    argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"] = xlWorkSheet.Cells[xlIterator, "AW"].Value.Trim().Substring(0,
                        xlWorkSheet.Cells[xlIterator, "AW"].Value.Trim().IndexOf(" "));
                    argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"] = arabNormalize(argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"].ToString());
                    term = " [семестр " + argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"].ToString() + "]";
                    //Add to table column "всього"
                    dsRobPlan.Tables[currentGroupName].Columns.Add(xlWorkSheet.Cells[16, "AS"].Value+ " годин" + term);
                    zaput[iZaput] = "AS";
                    iZaput++;

                    //Add to table column "курсові роботи, проекти"
                    dsRobPlan.Tables[currentGroupName].Columns.Add("КП" + term);
                    zaput[iZaput] = "BE";
                    iZaput++;

                    //Add to table column "екзамен"
                    dsRobPlan.Tables[currentGroupName].Columns.Add("Іспит" + term);
                    zaput[iZaput] = "BI";
                    iZaput++;

                    //Add to table column "диф  залік" or "залік"
                    s = xlWorkSheet.Cells[18, "BK"].Value.ToString();
                    if (s.Trim().Equals("диф  залік")) s2 = "Д/З";
                    else s2 = "Залік";
                    dsRobPlan.Tables[currentGroupName].Columns.Add(s2 + term);
                    zaput[iZaput] = "BK";
                    iZaput++;

                    //Add to table column "диф  залік" - if exist
                    if (countFormControl == 3)
                    {
                        dsRobPlan.Tables[currentGroupName].Columns.Add("Д/З" + term);
                        zaput[iZaput] = "BL";
                        iZaput++;
                    }

                    //Add to table collumn "Циклова комісія,\nвикладач"
                    if (xlWorkSheet.Cells[15, "BN"].Value.Equals("Циклова комісія,\nвикладач"))
                    {
                        dsRobPlan.Tables[currentGroupName].Columns.Add("викладач");
                        zaput[iZaput] = "BN";
                        iZaput++;
                    }
                    continue;
                }
                DataRow dataRow = dsRobPlan.Tables[currentGroupName].NewRow();

                //Записування усіх даних у нашу таблицю
                for (int i = 0; i < dsRobPlan.Tables[currentGroupName].Columns.Count; i++)
                {
                    var whatType = xlWorkSheet.Cells[xlIterator, zaput[i]].Value;
                    if (whatType != null) dataRow[i] = whatType;
                    else dataRow[i] = 0;

                    //Збільшення значень у полу 'Іспит' на 100, якщо це є державний екзамен
                    if(countryPas.Rows.Count > 0)
                    if (countryPas.Rows[0][0] != null && (dataRow.Table.Columns[i].ColumnName.ToString().Contains("Назви") ||
                        dataRow.Table.Columns[i].ColumnName.ToString().Contains("Іспит")) )
                    {
                        foreach (DataRow dr in countryPas.Rows)
                        {
                            if (dr[0].ToString().ToLower().Contains(dataRow[0].ToString().ToLower()) 
                                && dataRow.Table.Columns[i].ColumnName.ToString().Equals("Іспит [семестр " + dr[1].ToString().Trim() + "]"))
                            {
                                dataRow[i] = Convert.ToDouble(dataRow[i].ToString()) + 100;
                            }
                        }
                    }
                }
                dsRobPlan.Tables[currentGroupName].Rows.Add(dataRow);
            }

            //Зчитування та записування у dsRobPlan практик
            // - семестр - назва практики - число годин - викладач
            String[] locColumn = { "B", "C", "AA", "AJ" };
            int[] locRow = { 40, 41, 43, 47, 49 };
            for (int i = 0; i < locRow.Length; i++)
            {
                var xl = xlWorkSheet.Cells[locRow[i], locColumn[1]].Value;
                if (xl != null)
                {
                    if (xl.ToString().Contains("Назва практики"))
                    {
                        while (true)
                        {
                            locRow[i]++;
                            var xl2 = xlWorkSheet.Cells[locRow[i], locColumn[1]].Value;
                            if (xl2 == null) break;
                            if (xl2.ToString().Contains("Навчальна") || xl2.ToString().Contains("Виробнича")) continue;
                            
                            DataRow newRow = dsRobPlan.Tables[currentGroupName].NewRow();
                            newRow[0] = xlWorkSheet.Cells[locRow[i], locColumn[1]].Value.ToString();

                            for (int j = 1; j < newRow.Table.Columns.Count; j++)
                            {
                                if (newRow.Table.Columns[j].ColumnName.ToString().Contains("всього годин [семестр " +
                                    arabNormalize(xlWorkSheet.Cells[locRow[i], locColumn[0]].Value.ToString()) + "]"))
                                {
                                    newRow[j] = xlWorkSheet.Cells[locRow[i], locColumn[2]].Value.ToString();
                                }
                                else if (newRow.Table.Columns[j].ColumnName.ToString().Contains("КП [семестр " +
                                            arabNormalize(xlWorkSheet.Cells[locRow[i], locColumn[0]].Value.ToString()) + "]"))
                                    newRow[j] = -1;
                                else
                                    newRow[j] = 0;
                                if (j == newRow.Table.Columns.Count - 1)
                                {
                                    newRow[j] = xlWorkSheet.Cells[locRow[i], locColumn[3]].Value.ToString();
                                    dsRobPlan.Tables[currentGroupName].Rows.Add(newRow);
                                }
                            }
                        }
                    }
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }

        public void LoadData_StudDB(String pathStudDB)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(pathStudDB);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            dsRobPlan.Tables.Add("Студенти");
            dsRobPlan.Tables["Студенти"].Columns.Add("пільги");
            dsRobPlan.Tables["Студенти"].Columns.Add("піб");
            dsRobPlan.Tables["Студенти"].Columns.Add("номер книги");
            dsRobPlan.Tables["Студенти"].Columns.Add("група");
            dsRobPlan.Tables["Студенти"].Columns.Add("форма");
            

            int xlIterator = 2;
            while (true)
            {
                xlIterator++;
                if (xlIterator == 2 && xlWorkSheet.Cells[xlIterator, "C"].Value == null)
                {
                    MessageBox.Show("Помилка!\nПри відкритті книги 'Студенти' в клітині C2 очікувалося Ім'я студента");
                    xlWorkBook.Close();
                    xlApp.Quit();
                    return;
                }
                if (xlWorkSheet.Cells[xlIterator, "C"].Value == null) break;

                DataRow dataRow = dsRobPlan.Tables["Студенти"].NewRow();
                String cells = null;
                for (int i = 1; i < 6; i++)
                {
                    switch (i)
                    {
                        case 1:
                            cells = "L";
                            break;
                        case 2:
                            cells = "C";
                            break;
                        case 3:
                            cells = "D";
                            break;
                        case 4: 
                            cells = "E";
                            break;
                        case 5:
                            cells = "G";
                            break;
                    }
                    if (xlWorkSheet.Cells[xlIterator, cells].Value == null)
                        dataRow[i - 1] = " ";
                    else dataRow[i - 1] = xlWorkSheet.Cells[xlIterator, cells].Value.ToString();
                }
                dsRobPlan.Tables["Студенти"].Rows.Add(dataRow);
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }

        private String numberOfOblic = null, subject = null, arabSemestr = null;
        private int semestr = 0;
        public void createOblicUspishnosti(String numberOfOblic, int semestr, String subject)
        {
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            this.numberOfOblic = numberOfOblic;
            this.semestr = semestr;
            this.subject = subject;

            if (semestr == 1) arabSemestr = argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"].ToString();
            else arabSemestr = argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"].ToString();

            foreach (DataRow dr in dsRobPlan.Tables[currentGroupName].Rows)
            {
                if (dr[0].ToString().Equals(subject))
                {
                    
                    bool bl = true;
                    foreach (DataColumn dc in dr.Table.Columns)
                    {
                        if (dc.ColumnName.ToString().Contains(" [семестр " + arabSemestr + "]")
                            && !dr[dc.ToString()].ToString().Equals("0")
                            && !dc.ColumnName.ToString().Contains("всього"))
                        {
                            bl = false;
                            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Облік Успішності\" + currentGroupName + ".xls";
                            if (System.IO.File.Exists(path))
                            {
                                xlWorkBook = xlApp.Workbooks.Open(path);
                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                                //Якщо у назві листа не міститься поточного семестру то така робоча книга перезаписується
                                if (!xlWorkSheet.Name.ToString().Contains("_" + arabSemestr + "_"))
                                {
                                    if (MessageBox.Show("Відомість успішності для групи " + currentGroupName +
                                        " буде перезаписана", "Увага!", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                                    {
                                        xlWorkBook.Close();
                                        xlApp.Quit();
                                        return;
                                    }

                                    xlWorkBook.Close();
                                    System.IO.File.Copy(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                            + @"Data\Empty book.xls", path, true);
                                    xlWorkBook = xlApp.Workbooks.Open(path);
                                }
                            }
                            else
                            {
                                System.IO.File.Copy(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                    + @"Data\Empty book.xls", path, true);
                                xlWorkBook = xlApp.Workbooks.Open(path);
                            }

                            //Якщо предмет у робочій книзі створено, то його перезаписується
                            bool ifExist = true;
                            for (int i = 1; i <= xlWorkBook.Sheets.Count; i++)
                            {
                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                                
                                if (xlWorkSheet.Name.Contains("Sheet") || xlWorkSheet.Name.Contains("Листок") || xlWorkSheet.Name.Contains("Лист")
                                    || xlWorkSheet.Name.Contains("Аркуш"))
                                {
                                    xlWorkSheet.Name = cutSheetName(subject, "_" + arabSemestr + "_");
                                    ifExist = false;
                                    xlWorkBook.Save();
                                    break;
                                }
                                else if (subject.Contains(xlWorkSheet.Name.Substring(0, xlWorkSheet.Name.IndexOf("_"))))
                                {
                                    //xlWorkSheet.Delete();
                                    //xlWorkBook.Save();

                                    ifExist = false;
                                    break;
                                }
                            }
                            if (ifExist)
                            {
                                xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                xlWorkBook.Worksheets.get_Item(1).Name = cutSheetName(subject, "_" + arabSemestr + "_");
                                xlWorkBook.Save();
                            }

                            String tamplatePath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9) + @"Data\";

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(cutSheetName(subject, "_" + arabSemestr + "_"));

                            //Визначення форми здачі та запущення відповідного механізму
                            if (dc.ColumnName.ToString().Contains("КП") || dr[dc.ToString()].ToString().Equals("-1"))
                            {
                                tamplatePath += "Відомість обліку успішності (КП - Технологічна практика).xls";
                                if (dr[dc.ToString()].ToString().Equals("-1")) driverToKP(tamplatePath, true);
                                else driverToKP(tamplatePath, false);
                            }
                            else if (dc.ColumnName.ToString().Contains("Іспит") && Convert.ToDouble(dr[dc.ToString()]) > 100)
                            {
                                tamplatePath += "Відомість обліку успішності (Державний екзамен) - протокол.xls";
                                DriverToDerzPas(tamplatePath);
                            }
                            else
                            {
                                tamplatePath += "Відомість обліку успішності (залік - диф. залік - екзамен).xls";
                                driverToOblicOfZalic(tamplatePath);
                            }
                            xlWorkBook.Close();
                        }
                    }
                    if (bl)
                    {
                        MessageBox.Show("У " + semestr.ToString() + " півріччі вказаний предмет відсутній\nВкажіть інше півріччя");
                        xlApp.Quit();
                        return;
                    }
                    break;
                }

            }
            
            xlApp.Quit();
        }

        private void reloadSheet(String path)
        {
            Excel.Workbook tamplateBook = xlApp.Workbooks.Open(path);
            Excel.Worksheet tamplateSheet = (Excel.Worksheet)tamplateBook.Worksheets.get_Item(1);

            String nameSheet = xlWorkSheet.Name;
            xlApp.Visible = true;

            //Переприсвоєння імені із видаленням листка
            tamplateSheet.Copy(xlWorkSheet);
            xlWorkBook.Save();

            xlWorkSheet.Application.DisplayAlerts = false;
            
            xlWorkSheet.Delete();
            xlWorkBook.Save();

            try
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(nameSheet);
            }
            catch (Exception ex)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Name = nameSheet;

                xlWorkBook.Save();
                tamplateBook.Close();
                return;
            }

            xlWorkSheet.Application.DisplayAlerts = false;
            xlWorkSheet.Delete();
            xlWorkBook.Save();
        }

        private void DriverToDerzPas(String path)
        {
            reloadSheet(path);

            xlWorkSheet.Cells[4, "H"].Value = subject;
            xlWorkSheet.Cells[9, "C"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Код спеціальності"].ToString() + "_" + currentGroupName;
            for (int i = 0; i < dsRobPlan.Tables[currentGroupName].Rows.Count; i++)
            {
                if (dsRobPlan.Tables[currentGroupName].Rows[i][0].Equals(subject))
                {
                    for (int j = 1; j < dsRobPlan.Tables[currentGroupName].Columns.Count; j++)
                    {
                        if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("викладач"))
                        {
                            //Викладач
                            xlWorkSheet.Cells[20, "G"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString() + "_________________________________";
                            xlWorkSheet.Cells[84, "H"].Value = "_____" + dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString() + "__";
                        }
                    }
                    xlWorkBook.Save();
                    break;
                }
            }
            int n = 46;
            foreach (DataRow dr in dsRobPlan.Tables["Студенти"].Rows)
            {
                if (dr[3].ToString().Equals(currentGroupName))
                {
                    xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                    n++;
                }
            }

            //Кількість студентів у групі
            xlWorkSheet.Cells[12, "G"] = "__" + (n - 46) + "__";

            xlWorkBook.Save();
            if (n != 76)
                xlWorkSheet.Range["B" + n, "Q" + 75].Delete();
            xlWorkBook.Save();
        }

        private void driverToKP(String path, bool practuca)
        {
            reloadSheet(path);

            //Відділення
            String viddilenia = null;
            if (argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString().Equals("Програмна інженерія")) viddilenia = "Програмної інженерії";
            else viddilenia = "Метрології та інформаційно-вимірювальної технології";
            xlWorkSheet.Cells[13, "E"].Value = viddilenia;

            //Спеціальність
            xlWorkSheet.Cells[15, "F"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"].ToString();

            //Курс
            xlWorkSheet.Cells[17, "D"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Курс"].ToString();

            //Група
            xlWorkSheet.Cells[17, "G"].Value = currentGroupName;

            //Навчальний рік
            int year = Convert.ToInt32(argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString()) + 1;
            xlWorkSheet.Cells[19, "I"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString() + "-" + year;

            //Назва дисципліни
            xlWorkSheet.Cells[26, "F"].Value = subject;

            //Семестр
            xlWorkSheet.Cells[28, "D"].Value = arabSemestr;

            //Номер відомості
            xlWorkSheet.Cells[22, "M"].Value = numberOfOblic.ToString();

            for (int i = 0; i < dsRobPlan.Tables[currentGroupName].Rows.Count; i++)
            {
                if (dsRobPlan.Tables[currentGroupName].Rows[i][0].Equals(subject))
                {
                    for (int j = 1; j < dsRobPlan.Tables[currentGroupName].Columns.Count; j++)
                    {
                        if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("[семестр " + arabSemestr + "]"))
                        {
                            if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("всього годин"))
                            {
                                //Всього годин
                                if (!dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString().Equals("0"))
                                xlWorkSheet.Cells[30, "Q"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString();
                            }
                            else if (!dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString().Equals("0"))
                            {
                                //Форма здачі
                                if (!practuca) xlWorkSheet.Cells[30, "F"].Value = "курсовий проект";
                                else
                                {
                                    if (subject.Trim().Equals("Технологічна практика") || subject.Trim().Equals("Переддипломна практика"))
                                        xlWorkSheet.Cells[30, "F"].Value = "захист";
                                    else
                                        xlWorkSheet.Cells[30, "F"].Value = "Д/З";
                                }
                            }
                        }
                        else if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("викладач"))
                        {
                            //Викладач
                            xlWorkSheet.Cells[37, "K"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString() + "_____";
                            xlWorkSheet.Cells[100, "N"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString();
                        }
                    }
                    xlWorkBook.Save();
                    break;
                }
            }
            int n = 45;
            foreach (DataRow dr in dsRobPlan.Tables["Студенти"].Rows)
            {
                if (dr[3].ToString().Equals(currentGroupName))
                {
                    xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                    xlWorkSheet.Cells[n, "H"].Value = dr[2].ToString();
                    n++;
                }
            }
            xlWorkBook.Save();
            if (n != 75)
                xlWorkSheet.Range["B" + n, "Q" + 74].Delete();
            xlWorkBook.Save();
        }

        private void driverToOblicOfZalic(String path)
        {
            reloadSheet(path);

            //Відділення
            String viddilenia = null;
            if (argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString().Equals("Програмна інженерія")) viddilenia = "Програмної інженерії";
            else viddilenia = "Метрології та інформаційно-вимірювальної технології";
            xlWorkSheet.Cells[13, "E"].Value = viddilenia;

            //Спеціальність
            xlWorkSheet.Cells[15, "F"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"].ToString();

            //Курс
            xlWorkSheet.Cells[17, "D"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Курс"].ToString();

            //Група
            xlWorkSheet.Cells[17, "G"].Value = currentGroupName;

            //Навчальний рік
            int year = Convert.ToInt32(argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString()) + 1;
            xlWorkSheet.Cells[19, "I"].Value = argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString() + "-" + year;

            //Назва дисципліни
            xlWorkSheet.Cells[26, "F"].Value = subject;

            //Семестр
            xlWorkSheet.Cells[28, "D"].Value = arabSemestr;

            //Номер відомості
            xlWorkSheet.Cells[22, "M"].Value = numberOfOblic.ToString();

            for (int i = 0; i < dsRobPlan.Tables[currentGroupName].Rows.Count; i++)
            {
                if (dsRobPlan.Tables[currentGroupName].Rows[i][0].Equals(subject))
                {
                    for (int j = 1; j < dsRobPlan.Tables[currentGroupName].Columns.Count; j++)
                    {
                        if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("[семестр " + arabSemestr + "]"))
                        {
                            if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("всього годин"))
                            {
                                //Всього годин
                                xlWorkSheet.Cells[30, "Q"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString();
                            }
                            else if (!dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString().Equals("0"))
                            {
                                //Форма здачі
                                xlWorkSheet.Cells[30, "F"].Value =
                                    dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Substring(0, dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.IndexOf(" "));
                            }
                        }
                        else if (dsRobPlan.Tables[currentGroupName].Columns[j].ColumnName.Contains("викладач"))
                        {
                            //Викладач
                            xlWorkSheet.Cells[32, "E"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString();
                            xlWorkSheet.Cells[94, "N"].Value = dsRobPlan.Tables[currentGroupName].Rows[i][j].ToString();
                        }
                    }
                    xlWorkBook.Save();
                    break;
                }
            }
            int n = 39;
            foreach (DataRow dr in dsRobPlan.Tables["Студенти"].Rows)
            {
                if (dr[3].ToString().Equals(currentGroupName))
                {
                    xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                    xlWorkSheet.Cells[n, "H"].Value = dr[2].ToString();
                    n++;
                }
            }
            xlWorkBook.Save();
            if (n != 69)
            xlWorkSheet.Range["B" + n, "Q" + 68].Delete();
            xlWorkBook.Save();
        }

        private String arabNormalize(String str)
        {
            char[] ch = str.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                int arg = (int)ch[i];
                if (arg == 1030) arg = 73;
                ch[i] = (char)arg;
            }
            return new String(ch);
        }

        private String cutSheetName(String s, String arab)
        {
            int lenght = s.Length + arab.Length;
            if (lenght <= 32){
                return (s.Replace("*","&") + arab);

            }
                
            else
                return (s.Substring(0, 31 - arab.Length).Replace("*", "&") + arab);
        }

        public DataTable dtSo = null;
        public void zvedVidomist(int semestr, String subject, String mount)
        {
            this.subject = subject;
            this.semestr = semestr;


            if (semestr == 1) arabSemestr = argDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"].ToString();
            else arabSemestr = argDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"].ToString();

            String sT = null;
            if (mount.Equals(""))
            {
                if (semestr == 1) sT = "1-ше півріччя";
                else sT = "2-ге півріччя";
            }
            else sT = mount + " місяць";

            String existsPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\" + "Empty book.xls";

            String pathTogroup = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Зведена відомість успішності\" + "Зведена відомість успішності за " + sT + ".xls";

            String pathToTamplateVidomist = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\" + "Зведена відомість.xls";

            if (!System.IO.File.Exists(pathTogroup))
                System.IO.File.Copy(existsPath, pathTogroup, true);

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(pathTogroup);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            bool ifExist = true;
            for (int i = 1; i <= xlWorkBook.Sheets.Count; i++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                if (xlWorkSheet.Name.Contains("Sheet") || xlWorkSheet.Name.Contains("Аркуш") || xlWorkSheet.Name.Contains("Лист"))
                {
                    xlWorkSheet.Name = currentGroupName;
                    ifExist = false;
                    xlWorkBook.Save(); 
                    break;
                }
                else if (xlWorkSheet.Name.Equals(currentGroupName))
                {
                    if (MessageBox.Show("Для поточної групи (" + currentGroupName + ") уже створено зведену відомість\n Перезаписати?", "Обережно, ви можете втратити дані", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        xlWorkBook.Close();
                        xlApp.Quit();
                        return;
                    }
                    //xlWorkSheet.Delete();
                    ifExist = false;
                    //xlWorkBook.Save();
                    break;
                }
            }
            if (ifExist)
            {
                xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlApp.DisplayAlerts = false;
                xlWorkBook.Worksheets.get_Item(1).Name = currentGroupName;
                xlWorkBook.Save();
            }
            xlWorkSheet = xlWorkBook.Worksheets.get_Item(currentGroupName);
            reloadSheet(pathToTamplateVidomist);

            String[] forma = { "Іспит", "Д/З", "КП", "Залік", "КП" };
            char startX = 'E';
            char pos = ++startX;
            startX--;
            int count = 0;
            int otherCount = 0;
            for (int i = 0; i < forma.Length; i++)
            {
                bool bl = false;
                for (int row = 0; row < dsRobPlan.Tables[currentGroupName].Rows.Count; row++)
                {
                    try
                    {
                        if (!dsRobPlan.Tables[currentGroupName].Rows[row][forma[i] + " [семестр " + arabSemestr + "]"].ToString().Equals("0") && i != 2 && i != 4)
                        {
                            bl = true;
                            startX++;
                            xlWorkSheet.Cells[9, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                            xlWorkSheet.Cells[43, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["викладач"].ToString();

                            xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[9, startX.ToString()].ColumnWidth = columnWidth(dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());

                            count++;
                        }
                    }
                    catch (ArgumentException ar)
                    {
                        continue;
                    }
                    if (row == dsRobPlan.Tables[currentGroupName].Rows.Count - 1 && i != 2 && i != 4 && bl)
                    {
                        xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Merge();
                        
                        xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Value = forma[i];

                        xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                        pos = startX;
                        pos++;

                        bl = false;
                    }

                    if (!mount.Equals("")) continue;

                    if (!dsRobPlan.Tables[currentGroupName].Rows[row][forma[i] + " [семестр " + arabSemestr + "]"].ToString().Equals("0") && i == 2 && i != 4
                        && !dsRobPlan.Tables[currentGroupName].Rows[row][forma[i] + " [семестр " + arabSemestr + "]"].ToString().Equals("-1"))
                    {
                        startX++;
                        xlWorkSheet.Cells[9, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                        xlWorkSheet.Cells[43, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["викладач"].ToString();

                        xlWorkSheet.Cells[9, startX.ToString()].HorizontalAlignment = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        xlWorkSheet.Cells[43, startX.ToString()].ColumnWidth = columnWidth(dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());

                        count++;
                        pos = startX;
                        pos++;
                    }

                    if (i == 4 && dsRobPlan.Tables[currentGroupName].Rows[row][forma[i] + " [семестр " + arabSemestr + "]"].ToString().Equals("-1"))
                    {
                        otherCount++;
                        startX++;
                        xlWorkSheet.Cells[9, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                        xlWorkSheet.Cells[43, startX.ToString()].Value = dsRobPlan.Tables[currentGroupName].Rows[row]["викладач"].ToString();

                        xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[43, startX.ToString()].ColumnWidth = columnWidth(dsRobPlan.Tables[currentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());
                        
                    }
                }
                xlWorkBook.Save();
            }
            //добавлення та формування клітини з середнім балом
            startX++;
            xlWorkSheet.Cells[9, startX.ToString()].Value = "Середній бал";
            xlWorkSheet.Cells[9, startX.ToString()].ColumnWidth = 6.2;
            xlWorkSheet.Cells[9, startX.ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            pos = 'E';
            for (int i = 0; i < count; i++)
                pos++;

            xlWorkSheet.Range["F7", pos.ToString() + 7].Merge();
            xlWorkSheet.Range["F7", pos.ToString() + 7].Value = "Предмети";
            xlWorkSheet.Range["F7", pos.ToString() + 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Range["F7", pos.ToString() + 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlWorkBook.Save();

            
            int startY = 11;
            int studCount = 0;
            foreach (DataRow dt in dsRobPlan.Tables["Студенти"].Rows)
            {
                if (dt[3].ToString().Equals(currentGroupName))
                {
                    xlWorkSheet.Range["D" + startY].Value = dt[1].ToString();
                    if (dt[4].ToString().Equals("п")) xlWorkSheet.Range["E" + startY].Value = dt[4].ToString();
                    startY++;
                    studCount++;
                }
            }

            //куратор
            xlWorkSheet.Range["K45"].Value = "/ " + currentCurator() + " /";

            if (startY != 40)
                xlWorkSheet.Range["A" + startY, "IV" + 40].Delete();

            startY += 3;
            xlWorkSheet.Range["C7", startX + startY.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            if (mount.Equals(""))
            {
                xlWorkSheet.Range["C4"].Value = "спеціальності \"" + argDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"].ToString() + "\"";
                int year = Convert.ToInt32(argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString()) + 1;
                xlWorkSheet.Range["D5"].Value = "групи " + currentGroupName + " за " + arabSemestr + " семестр " +
                argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString() + "-" + year.ToString() + " навчального року";
            }
            else
            {
                xlWorkSheet.Range["D4"].Value = "спеціальності \"" + argDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"].ToString() + "\"";
                int year = Convert.ToInt32(argDataSet.Tables[currentGroupName].Rows[0]["Рік"].ToString()) + 1;
                xlWorkSheet.Range["C5"].Value = "за місяць " + mount + year + "р.";
                xlWorkSheet.Range["C6"].Value = "група " + currentGroupName;
            }
            
            String pathToOblic  = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Облік успішності\" + currentGroupName + ".xls";

            xlWorkBook.Save();
            //занесення оцінок
            if (System.IO.File.Exists(pathToOblic) && mount.Equals(""))
            {
                //прохід через усі записані у зведену відомість предмети
                DataTable dataTable = getThePas(pathToOblic);

                startX = 'E';
                for (int i = 0; i < count + otherCount; i++)
                {
                    startX++;
                    String likeSheet = cutSheetName(xlWorkSheet.Range[startX.ToString() + 9].Value.ToString(), "_" + arabSemestr + "_");
                    //MessageBox.Show(likeSheet);
                    int negatPasCount = 0, superNegativPasCount = 0;
                    int currentRow = 0;
                    for (int j = 1; j <= studCount; j++)
                    {
                        try
                        {
                            bool bl = dataTable.Rows[j - 1][likeSheet].ToString().Equals("0");
                            if (bl) continue;
                        }
                        catch (Exception e)
                        {
                            break;
                        }
                        currentRow++;
                        String pas = dataTable.Rows[currentRow - 1][likeSheet].ToString();
                        int cell = j + 10;
                        xlWorkSheet.Range[startX.ToString() + cell].Value = pas;
                        if (pas == null || pas.Equals("") || pas.Equals(" "))
                        {
                            continue;
                        }

                        if (!pas.Contains("з"))
                        {
                            try
                            {
                                int number = Convert.ToInt32(pas);
                                if (number <= 3) superNegativPasCount++;
                                else if (number < 7) negatPasCount++;
                            }
                            catch (Exception ex) { }
                        }
                    }
                    xlWorkSheet.Range[startX.ToString() + (studCount + 11)].Formula = "=" + (studCount - superNegativPasCount) + "/" + studCount;
                    xlWorkSheet.Range[startX.ToString() + (studCount + 12)].Formula = "=" + (studCount - (superNegativPasCount + negatPasCount)) + "/" + studCount;
                }
                xlWorkBook.Save();
                startX = 'E';

                for (int i = 0; i < count + otherCount; i++)
                    startX++;

                char begin = startX;
                char averageBal = ++startX;
                startX++;
                
                for (int i = 11; i < studCount + 11; i++)
                {
                    xlWorkSheet.Range[averageBal.ToString() + i].Formula = "=AVERAGE(" + "F" + i + ":" + begin.ToString() + i + ")";
                    xlWorkSheet.Range[averageBal.ToString() + i].NumberFormatLocal = "##,##";
                    String s1 = xlWorkSheet.Range["E" + i].Value;
                    Double s2;
                    try
                    {
                        s2 = xlWorkSheet.Range[averageBal.ToString() + i].Value;
                    }
                    catch (Exception ex) { continue; }

                    if (s1 == null)
                    {
                        if (s2 > 7)
                        {
                            xlWorkSheet.Range[startX.ToString() + i].Value = "1";
                            if (dataTable.Rows[i - 11]["підвищена стипендія"].ToString().Equals("1"))
                                xlWorkSheet.Range[startX.ToString() + i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }
                }
                xlWorkBook.Save();
            }

            startX = 'E';

            for (int i = 0; i < count + otherCount + 1; i++)
                startX++;

            char stypendia = ++startX;
            startX++;
            int xlIter = 10;
            foreach (DataRow dt in dsRobPlan.Tables["Студенти"].Rows)
            {
                if (dt[3].ToString().Equals(currentGroupName))
                {
                    xlIter++;
                    var whatType = dt[0];
                    if (whatType != null)
                    {
                        bool bl = true;
                        var whatype2 = xlWorkSheet.Range[stypendia.ToString() + xlIter].Value;
                        if (whatype2 == null) bl = false;

                        if (whatType.ToString().Contains("сир"))
                            xlWorkSheet.Range[startX.ToString() + xlIter].Value = "с";
                        else if (whatType.ToString().Contains("гір") && bl)
                            xlWorkSheet.Range[startX.ToString() + xlIter].Value = "г";
                        else if (whatType.ToString().Contains("інва"))
                            xlWorkSheet.Range[startX.ToString() + xlIter].Value = "і";
                    }
                }
            }
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();
        }

        private DataTable getThePas(String path)
        {
            DataTable dt = new System.Data.DataTable();

            Excel.Workbook newBook = xlApp.Workbooks.Open(path);
            Excel.Worksheet newSheet = null;

            dt.Columns.Add("ID", typeof(Int32));
            dt.Columns[0].AllowDBNull = false;
            dt.Columns[0].AutoIncrement = true;
            dt.Columns[0].AutoIncrementStep = 1;
            dt.Columns[0].Unique = true;

            for (int i = 1; i <= newBook.Worksheets.Count; i++)
            {
                dt.Columns.Add(newBook.Worksheets.get_Item(i).Name.ToString(), typeof(String));
            }
            
            String [] st = {"B", "J", "46", "B", "L", "39", "B", "L", "45"};
            int pos = 0;
            for (int i = 1; i <= newBook.Worksheets.Count; i++)
            {
                newSheet = (Excel.Worksheet)newBook.Worksheets.get_Item(i);
                for (int j = 0; j < st.Length; j += 3)
                {
                    var whatType = newSheet.Range[st[j] + st[j + 2]].Value;
                    if (whatType != null)
                    {
                        if (whatType.Equals("1."))
                        { 
                            pos = j;
                            break;
                        }
                    }
                }

                for (int j = Convert.ToInt32(st[pos + 2]); ; j++)
                {
                    if (i == 1)
                    {
                        DataRow dataRow = dt.NewRow();
                        dt.Rows.Add(dataRow);
                    }
                    var whatType = newSheet.Range[st[pos] + j].Value;
                    if (whatType == null) break;
                    var ocinka = newSheet.Range[st[pos + 1] + j].Value;
                    if (ocinka == null) ocinka = " ";
                    dt.Rows[j - Convert.ToInt32(st[pos + 2])][newSheet.Name.ToString()] = ocinka.ToString();
                }
            }

            dt.Columns.Add("підвищена стипендія");
            for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                bool ifHight = true;
                dt.Rows[i]["підвищена стипендія"] = "0";
                for (int j = 1; j < dt.Columns.Count - 2; j++)
                {
                    try
                    {
                        if (Convert.ToInt32(dt.Rows[i][j]) < 10)
                            ifHight = false;
                    }
                    catch (Exception ex) {  }
                }

                if (ifHight)
                    dt.Rows[i]["підвищена стипендія"] = "1";
            }

            newBook.Close();
            return dt;
        }

        private double columnWidth(String s)
        {
            if (s.Length <= 21) return 5.57;
            else if (s.Length <= 40) return 9.70;
            else if (s.Length <= 55) return 11;
            else return 13.43;
        }

        public void ArhiveZvedVid(String path)
        {
            if (!path.Contains("Зведена відомість успішності за"))
            {
                MessageBox.Show("Помилка!\nНазви excel-файлів повинні співпадати із назвами заданими у програмі");
                return;
            }

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(path);
            
            String groupArxivePath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Зведена відомість успішності\Архів\";

            String emptyBookPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\Empty book.xls";

            for (int i = 1; i <= xlWorkBook.Worksheets.Count; i++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                String sheetName = xlWorkSheet.Name;

                if (!System.IO.File.Exists(groupArxivePath + sheetName + ".xls"))
                {
                    int semestCount = 6;
                    if (sheetName.Contains("ПІ")) semestCount = 8;

                    System.IO.File.Copy(emptyBookPath, groupArxivePath + sheetName + ".xls", true);

                    Excel.Workbook bookNew = xlApp.Workbooks.Open(groupArxivePath + sheetName + ".xls");

                    for (int j = 1; j <= semestCount; j++)
                    {
                        bookNew.Worksheets.Add();
                        bookNew.Save();
                        bookNew.Worksheets.get_Item(1).Name = "семестр " + arabToRome(j);
                        bookNew.Save();
                    }
                    bookNew.Close();

                }
                reloadSheet2(groupArxivePath + sheetName + ".xls");
            }
        }

        private void reloadSheet2(String path)
        {
            Excel.Workbook bookNew = xlApp.Workbooks.Open(path);

            for (int j = 1; j <= bookNew.Worksheets.Count - 1; j++)
            {
                Excel.Worksheet sheetNew = (Excel.Worksheet)bookNew.Worksheets.get_Item(j);

                String semestr = xlWorkSheet.Range["D5"].Value.ToString().Substring(18);
                semestr = semestr.Substring(0, semestr.IndexOf(" "));
                //MessageBox.Show(">" + semestr + "<\n");

                String sheet = sheetNew.Name.Substring(sheetNew.Name.IndexOf(" ") + 1);
                //semestr.Substring(0, semestr.IndexOf(" "));

                //MessageBox.Show(">" + sheet + "<\n" + ">" + semestr + "<");
                if (semestr.Equals(sheet))
                {
                    //Переприсвоєння імені із видаленням листка
                    xlWorkSheet.Copy(sheetNew);
                    bookNew.Save();
                    String name = sheetNew.Name;

                    sheetNew.Application.DisplayAlerts = false;
                    sheetNew.Delete();
                    bookNew.Save();

                    try
                    {
                        sheetNew = (Excel.Worksheet)bookNew.Worksheets.get_Item(name);
                    }
                    catch (Exception ex)
                    {
                        sheetNew = (Excel.Worksheet)bookNew.Worksheets.get_Item(xlWorkSheet.Name);
                        sheetNew.Name = name;

                        bookNew.Save();
                    }
                }
            }
            bookNew.Close();
        }

        public String arabToRome(int arab)
        {
            String rome = "";
            switch (arab)
            {
                case 1:
                    rome = "I"; break;
                case 2:
                    rome = "II"; break;
                case 3:
                    rome = "III"; break;
                case 4:
                    rome = "IV"; break;
                case 5:
                    rome = "V"; break;
                case 6:
                    rome = "VI"; break;
                case 7:
                    rome = "VII"; break;
                case 8:
                    rome = "VIII"; break;

            }
            return rome;
        }

        public String currentCurator()
        {
            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                 + @"Data\Куратори.xls";

            Excel.Workbook book = xlApp.Workbooks.Open(path);
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets.get_Item("Куратори");

            for (int i = 2; ; i++)
            {
                if (sheet.Range["A" + i].Value == null) break;
                else if (sheet.Range["A" + i].Value.ToString().Equals(currentGroupName))
                {
                    return sheet.Range["B" + i].Value.ToString();
                }
            }

            book.Close();
            return "";
        }
    }
}