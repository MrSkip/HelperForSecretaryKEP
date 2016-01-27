using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using myKR.Properties;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public class ExcelWork
    {
        private Excel.Application _xlApp;
        private Excel.Workbook _xlWorkBook;
        private Excel.Worksheet _xlWorkSheet;

        //Path to Excel file
        private readonly string _pathRobPlan;

        // where argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString() - "Напрям підготовки", .[1] - "Спеціальність", .[2] - "Курс", .[3] - "рік"
        // .[4] - перше півріччя, .[5] - друге півріччя, .[6] - код спеціальності
        public DataSet ArgDataSet = new DataSet();
        public string[] RobPlanArgs = new string[7];

        //Group names of sheets from excel file "PlanRob"
        public string[] SheetNamesRobPlan = new string[20];

        //in this DataSet we load data from robBlan and StudDB
        public DataSet DsRobPlan = new DataSet();

        //
        public string CurrentGroupName;

        private string _numberOfOblic, _subject, _arabSemestr;

        public ExcelWork(string pathRobPlan)
        {
            _pathRobPlan = pathRobPlan;
            LoadSheetName_RobPlan();
        }

        public ExcelWork()
        {
        }

        public int Semestr { get; set; }

        private void LoadSheetName_RobPlan()
        {
            _xlApp = new Excel.Application {DisplayAlerts = false};
            _xlWorkBook = _xlApp.Workbooks.Open(_pathRobPlan);

            int countGroup = 0;
            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++)
            {
                string name = _xlWorkBook.Worksheets.get_Item(i).Name;
                if (name.Length != 8 || name.IndexOf('-') != 2) continue;
                SheetNamesRobPlan[countGroup] = _xlWorkBook.Worksheets.get_Item(i).Name;
                countGroup++;
            }

            _xlWorkBook.Close();
            _xlApp.Quit();
        }

        //Read need argument for our program from Ecxel file
        public void LoadData_RobPlan(string currentGroupName)
        {
            _xlApp = new Excel.Application();
            _xlWorkBook = _xlApp.Workbooks.Open(_pathRobPlan);

            CurrentGroupName = currentGroupName;
            _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[currentGroupName];

            // where argDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"].ToString() - "Напрям підготовки", .[1] - "Спеціальність", .[2] - "Курс", .[3] - "рік"
            // .[4] - перше півріччя, .[5] - друге півріччя, .[6] - код спеціальності
            try
            {
                ArgDataSet.Tables.Add(currentGroupName);
            }
            catch (Exception)
            {
                _xlWorkBook.Close();
                _xlApp.Quit();
                return;
            }

            ArgDataSet.Tables[currentGroupName].Columns.Add("Напрям підготовки");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Спеціальність");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Курс");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Рік");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Перше півріччя");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Друге півріччя");
            ArgDataSet.Tables[currentGroupName].Columns.Add("Код спеціальності");


            ArgDataSet.Tables[currentGroupName].Rows.Add(ArgDataSet.Tables[currentGroupName].NewRow());

            //Read "Напряму підготовки"
            string arg = _xlWorkSheet.Cells[6, "R"].Value.ToString().Trim();
            ArgDataSet.Tables[currentGroupName].Rows[0]["Напрям підготовки"] = arg.Substring(arg.IndexOf("\"", StringComparison.Ordinal) + 1, arg.LastIndexOf("\"", StringComparison.Ordinal) - arg.IndexOf("\"", StringComparison.Ordinal) - 1);

            //Read "Спеціальність"
            arg = _xlWorkSheet.Cells[7, "R"].Value;
            ArgDataSet.Tables[currentGroupName].Rows[0]["Спеціальність"] = arg.Substring(arg.IndexOf("\"", StringComparison.Ordinal) + 1, arg.LastIndexOf("\"", StringComparison.Ordinal) - arg.IndexOf("\"", StringComparison.Ordinal) - 1);
            
            //Код спеціальності
            arg = arg.Trim().Substring(0, arg.Trim().IndexOf(" \"", StringComparison.Ordinal));
            ArgDataSet.Tables[currentGroupName].Rows[0]["Код спеціальності"] = arg.Substring(arg.IndexOf(" ", StringComparison.Ordinal));

            //read "Курс"
            arg = _xlWorkSheet.Cells[9 , "R"].Value.Trim();
            arg = arg.Substring(7);
            ArgDataSet.Tables[currentGroupName].Rows[0]["Курс"] = arg.Remove(arg.IndexOf("_", StringComparison.Ordinal));

            //Read "Рік"
            ArgDataSet.Tables[currentGroupName].Rows[0]["Рік"] = _xlWorkSheet.Cells[6, "B"].Value.ToString().Substring(_xlWorkSheet.Cells[6, "B"].Value.ToString().Length - 9, 4);

            
            //Read need values from `RobPlan` to our `dsTable`
            int xlIterator = 14;
            int countFormControl = 2;

            string[] zaput = new string[12];
            int iZaput = 0;

            DataTable countryPas = new DataTable();
            countryPas.Columns.Add("pas");
            countryPas.Columns.Add("term");

            while (true)
            {
                xlIterator++;

                String value = _xlWorkSheet.Cells[xlIterator, "C"].Value;
                if (value == null) continue;
                if (!value.Equals("Назви навчальних  дисциплін") && xlIterator == 15)
                {
                    MessageBox.Show(string.Format("Назви значень полів у:\n{0}\nне співпадає із заданами значеннями у програмі\n" + "Джерело - {1}\n У клітині \"С15\" очікувалося значення 'Назви навчальних  дисциплін'", _pathRobPlan, currentGroupName));
                    _xlWorkBook.Close();
                    _xlApp.Quit();
                    return;
                }
                if (value.Trim().Equals("Разом")) break;

                if (xlIterator == 15)
                {
                    //Задаємо усі можиливі місця розміщення клітини з назвою про державний екзамен  - "Назва"
                    //[][0] - колонка з назвою екзамена, [][1] - колонка з рядком езамену, [][2] - колонка з назвою семестра проходження екзамену
                    //відповідно, назва екзамену і назва семестру будуть знаходитися у одному рядку
                    string[][] doubleCell = {
                                                new[] {"BE", "49", "BO"},
                                                new[] {"BE", "47", "BO"},
                                                new[] {"BE", "42", "BO"},
                                                new[] {"BE", "43", "BO"},
                                                new[] {"AX", "39", "BC"}
                                            };
                    //Цикл для проходження усіх можливих місць знаходження назви державного екзамену, клітини із значенням "Назва"
                    foreach (string[] t in doubleCell)
                    {
//Для унеможливлення винекнинне помилок автоматично присвоюємо тип змінній howType
                        //Якщо у нашій клітині є якесь значення, то виконується умова
                        var howType = _xlWorkSheet.Cells[t[1], t[0]].Value;
                        if (howType == null) continue;
                        //Перевіряємо чи це справді клітина з назвою про державний екзамен
                        if (!howType.ToString().Trim().Equals("Назва")) continue;
                        //Назви полів залишатимуться сталимим, натомість рядок постійно інкрементуватиметься, тому переводиму його до типу int
                        var iter = Convert.ToInt32(t[1]) + 1;
                        //Коли знайдено місце знаходження потрібної нам клітини, ми шукаємо чи є хоча б якісь екзамени
                        while (true)
                        {
                            var whatType = _xlWorkSheet.Cells[iter, t[0]].Value;
                            if (whatType != null)
                            {
                                //Якщо екзамен знайдено то записуємо у таблицю його назву і семестр у якому проводитиметься екзамен
                                countryPas.Rows.Add(whatType,
                                    ArabNormalize(_xlWorkSheet.Cells[iter, t[2]].Value.ToString().Trim()));
                                iter++;
                            }
                            else break;
                        }
                        break;
                    }
                    //Зчитування практики, якщо практика існує для поточного семестру то заносимо її у масив Mas
                    //Create table "currentTable"
                    DsRobPlan.Tables.Add(currentGroupName);

                    //Add to table column "Назви навчальних  дисциплін"
                    DsRobPlan.Tables[currentGroupName].Columns.Add(value);
                    zaput[iZaput] = "C";
                    iZaput++;

                    

                    //непарний семестр
                    ArgDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"] = _xlWorkSheet.Cells[xlIterator, "Y"].Value.ToString().Substring(0,
                        _xlWorkSheet.Cells[xlIterator, "Y"].Value.ToString().IndexOf(" "));
                    ArgDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"] = ArabNormalize(ArgDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"].ToString());

                    //Create string name of current term
                    var term = " [семестр " + ArgDataSet.Tables[currentGroupName].Rows[0]["Перше півріччя"] + "]";

                    //Add to table column "всього годин"
                    DsRobPlan.Tables[currentGroupName].Columns.Add(_xlWorkSheet.Cells[16, "Y"].Value + " годин" + term);
                    zaput[iZaput] = "Y";
                    iZaput++;

                    //Add to table column "курсові роботи, проекти"
                    DsRobPlan.Tables[currentGroupName].Columns.Add("КП" + term);
                    zaput[iZaput] = "AK";
                    iZaput++;

                    //Add to table column "екзамен"
                    DsRobPlan.Tables[currentGroupName].Columns.Add("Іспит" + term);
                    zaput[iZaput] = "AO";
                    iZaput++;

                    //Add to table column "диф  залік" or "залік"
                    string s = _xlWorkSheet.Cells[18, "AQ"].Value.ToString();
                    var s2 = s.Trim().Equals("диф  залік") ? "Д/З" : "Залік";
                    DsRobPlan.Tables[currentGroupName].Columns.Add(s2 + term);
                    zaput[iZaput] = "AQ";
                    iZaput++;

                    //Add to table column "диф  залік" - if exist
                    if (_xlWorkSheet.Cells[18, "AR"].Value != null)
                    {
                        DsRobPlan.Tables[currentGroupName].Columns.Add("Д/З" + term);
                        countFormControl = 3;
                        zaput[iZaput] = "AR";
                        iZaput++;
                    }

                    //парний семестр
                    ArgDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"] = _xlWorkSheet.Cells[xlIterator, "AW"].Value.Trim().Substring(0,
                        _xlWorkSheet.Cells[xlIterator, "AW"].Value.Trim().IndexOf(" "));
                    ArgDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"] = ArabNormalize(ArgDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"].ToString());
                    term = " [семестр " + ArgDataSet.Tables[currentGroupName].Rows[0]["Друге півріччя"] + "]";
                    //Add to table column "всього"
                    DsRobPlan.Tables[currentGroupName].Columns.Add(_xlWorkSheet.Cells[16, "AS"].Value+ " годин" + term);
                    zaput[iZaput] = "AS";
                    iZaput++;

                    //Add to table column "курсові роботи, проекти"
                    DsRobPlan.Tables[currentGroupName].Columns.Add("КП" + term);
                    zaput[iZaput] = "BE";
                    iZaput++;

                    //Add to table column "екзамен"
                    DsRobPlan.Tables[currentGroupName].Columns.Add("Іспит" + term);
                    zaput[iZaput] = "BI";
                    iZaput++;

                    //Add to table column "диф  залік" or "залік"
                    s = _xlWorkSheet.Cells[18, "BK"].Value.ToString();
                    s2 = s.Trim().Equals("диф  залік") ? "Д/З" : "Залік";
                    DsRobPlan.Tables[currentGroupName].Columns.Add(s2 + term);
                    zaput[iZaput] = "BK";
                    iZaput++;

                    //Add to table column "диф  залік" - if exist
                    if (countFormControl == 3)
                    {
                        DsRobPlan.Tables[currentGroupName].Columns.Add("Д/З" + term);
                        zaput[iZaput] = "BL";
                        iZaput++;
                    }

                    //Add to table collumn "Циклова комісія,\nвикладач"
                    if (_xlWorkSheet.Cells[15, "BN"].Value.Equals("Циклова комісія,\nвикладач"))
                    {
                        DsRobPlan.Tables[currentGroupName].Columns.Add("викладач");
                        zaput[iZaput] = "BN";
                        iZaput++;
                    }
                    continue;
                }
                DataRow dataRow = DsRobPlan.Tables[currentGroupName].NewRow();

                //Записування усіх даних у нашу таблицю
                for (int i = 0; i < DsRobPlan.Tables[currentGroupName].Columns.Count; i++)
                {
                    var whatType = _xlWorkSheet.Cells[xlIterator, zaput[i]].Value;
                    if (whatType != null) dataRow[i] = whatType;
                    else dataRow[i] = 0;

                    //Збільшення значень у полу 'Іспит' на 100, якщо це є державний екзамен
                    if (countryPas.Rows.Count <= 0) continue;
                    if (countryPas.Rows[0][0] == null ||
                        (!dataRow.Table.Columns[i].ColumnName.Contains("Назви") &&
                         !dataRow.Table.Columns[i].ColumnName.Contains("Іспит"))) continue;
                    foreach (DataRow dr in countryPas.Rows)
                    {
                        if (dr[0].ToString().ToLower().Contains(dataRow[0].ToString().ToLower()) 
                            && dataRow.Table.Columns[i].ColumnName.Equals("Іспит [семестр " + dr[1].ToString().Trim() + "]"))
                        {
                            dataRow[i] = Convert.ToDouble(dataRow[i].ToString()) + 100;
                        }
                    }
                }
                DsRobPlan.Tables[currentGroupName].Rows.Add(dataRow);
            }

            //Зчитування та записування у dsRobPlan практик
            // - семестр - назва практики - число годин - викладач
            String[] locColumn = { "B", "C", "AA", "AJ" };
            int[] locRow = { 40, 41, 43, 47, 49 };
            for (int i = 0; i < locRow.Length; i++)
            {
                var xl = _xlWorkSheet.Cells[locRow[i], locColumn[1]].Value;
                if (xl != null)
                {
                    if (xl.ToString().Contains("Назва практики"))
                    {
                        while (true)
                        {
                            locRow[i]++;
                            var xl2 = _xlWorkSheet.Cells[locRow[i], locColumn[1]].Value;
                            if (xl2 == null) break;
                            if (xl2.ToString().Contains("Навчальна") || xl2.ToString().Contains("Виробнича")) continue;
                            
                            DataRow newRow = DsRobPlan.Tables[currentGroupName].NewRow();
                            newRow[0] = _xlWorkSheet.Cells[locRow[i], locColumn[1]].Value.ToString();

                            for (int j = 1; j < newRow.Table.Columns.Count; j++)
                            {
                                if (newRow.Table.Columns[j].ColumnName.Contains("всього годин [семестр " +
                                    ArabNormalize(_xlWorkSheet.Cells[locRow[i], locColumn[0]].Value.ToString()) + "]"))
                                {
                                    newRow[j] = _xlWorkSheet.Cells[locRow[i], locColumn[2]].Value.ToString();
                                }
                                else if (newRow.Table.Columns[j].ColumnName.Contains("КП [семестр " +
                                            ArabNormalize(_xlWorkSheet.Cells[locRow[i], locColumn[0]].Value.ToString()) + "]"))
                                    newRow[j] = -1;
                                else
                                    newRow[j] = 0;
                                if (j != newRow.Table.Columns.Count - 1) continue;
                                newRow[j] = _xlWorkSheet.Cells[locRow[i], locColumn[3]].Value.ToString();
                                DsRobPlan.Tables[currentGroupName].Rows.Add(newRow);
                            }
                        }
                    }
                }
            }

            _xlWorkBook.Close();
            _xlApp.Quit();
        }

        public void LoadData_StudDB(String pathStudDb)
        {
            _xlApp = new Excel.Application();
            _xlWorkBook = _xlApp.Workbooks.Open(pathStudDb);
            _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[1];

            DsRobPlan.Tables.Add("Студенти");
            DsRobPlan.Tables["Студенти"].Columns.Add("пільги");
            DsRobPlan.Tables["Студенти"].Columns.Add("піб");
            DsRobPlan.Tables["Студенти"].Columns.Add("номер книги");
            DsRobPlan.Tables["Студенти"].Columns.Add("група");
            DsRobPlan.Tables["Студенти"].Columns.Add("форма");
                

            int xlIterator = 2;
            while (true)
            {
                xlIterator++;
                if (xlIterator == 2 && _xlWorkSheet.Cells[xlIterator, "C"].Value == null)
                {
                    MessageBox.Show(Resources.ExcelWork_LoadData_StudDB_);
                    _xlWorkBook.Close();
                    _xlApp.Quit();
                    return;
                }
                if (_xlWorkSheet.Cells[xlIterator, "C"].Value == null) break;

                var dataRow = DsRobPlan.Tables["Студенти"].NewRow();
                string cells = null;
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
                    if (_xlWorkSheet.Cells[xlIterator, cells].Value == null)
                        dataRow[i - 1] = " ";
                    else dataRow[i - 1] = _xlWorkSheet.Cells[xlIterator, cells].Value.ToString();
                }
                DsRobPlan.Tables["Студенти"].Rows.Add(dataRow);
            }

            _xlWorkBook.Close();
            _xlApp.Quit();
        }
        public void CreateOblicUspishnosti(string numberOfOblic, int semestr, string subject)
        {
            if (numberOfOblic == null) throw new ArgumentNullException("numberOfOblic");
            _xlApp = new Excel.Application();
            _xlApp.DisplayAlerts = false;

            _numberOfOblic = numberOfOblic;
            Semestr = semestr;
            _subject = subject;

            if (semestr == 1) _arabSemestr = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Перше півріччя"].ToString();
            else _arabSemestr = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Друге півріччя"].ToString();

            foreach (DataRow dr in DsRobPlan.Tables[CurrentGroupName].Rows)
            {
                if (dr[0].ToString().Equals(subject))
                {
                    
                    bool bl = true;
                    foreach (DataColumn dc in dr.Table.Columns)
                    {
                        if (dc.ColumnName.Contains(" [семестр " + _arabSemestr + "]")
                            && !dr[dc.ToString()].ToString().Equals("0")
                            && !dc.ColumnName.Contains("всього"))
                        {
                            bl = false;
                            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Облік Успішності\" + CurrentGroupName + ".xls";
                            if (System.IO.File.Exists(path))
                            {
                                _xlWorkBook = _xlApp.Workbooks.Open(path);
                                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[1];

                                //Якщо у назві листа не міститься поточного семестру то така робоча книга перезаписується
                                if (!_xlWorkSheet.Name.Contains("_" + _arabSemestr + "_"))
                                {
                                    if (MessageBox.Show("Відомість успішності для групи " + CurrentGroupName +
                                        " буде перезаписана", "Увага!", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                                    {
                                        _xlWorkBook.Close();
                                        _xlApp.Quit();
                                        return;
                                    }

                                    _xlWorkBook.Close();
                                    System.IO.File.Copy(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                            + @"Data\Empty book.xls", path, true);
                                    _xlWorkBook = _xlApp.Workbooks.Open(path);
                                }
                            }
                            else
                            {
                                System.IO.File.Copy(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                    + @"Data\Empty book.xls", path, true);
                                _xlWorkBook = _xlApp.Workbooks.Open(path);
                            }

                            //Якщо предмет у робочій книзі створено, то його перезаписується
                            bool ifExist = true;
                            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++)
                            {
                                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[i];
                                
                                if (_xlWorkSheet.Name.Contains("Sheet") || _xlWorkSheet.Name.Contains("Листок") || _xlWorkSheet.Name.Contains("Лист")
                                    || _xlWorkSheet.Name.Contains("Аркуш"))
                                {
                                    _xlWorkSheet.Name = CutSheetName(subject, "_" + _arabSemestr + "_");
                                    ifExist = false;
                                    _xlWorkBook.Save();
                                    break;
                                }
                                if (!subject.Contains(_xlWorkSheet.Name.Substring(0, _xlWorkSheet.Name.IndexOf("_", StringComparison.Ordinal))))
                                    continue;
                                //xlWorkSheet.Delete();
                                //xlWorkBook.Save();

                                ifExist = false;
                                break;
                            }
                            if (ifExist)
                            {
                                _xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                _xlWorkBook.Worksheets.get_Item(1).Name = CutSheetName(subject, "_" + _arabSemestr + "_");
                                _xlWorkBook.Save();
                            }

                            String tamplatePath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9) + @"Data\";

                            _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[CutSheetName(subject, "_" + _arabSemestr + "_")];

                            //Визначення форми здачі та запущення відповідного механізму
                            if (dc.ColumnName.Contains("КП") || dr[dc.ToString()].ToString().Equals("-1"))
                            {
                                tamplatePath += "Відомість обліку успішності (КП - Технологічна практика).xls";
                                if (dr[dc.ToString()].ToString().Equals("-1")) DriverToKp(tamplatePath, true);
                                else DriverToKp(tamplatePath, false);
                            }
                            else if (dc.ColumnName.Contains("Іспит") && Convert.ToDouble(dr[dc.ToString()]) > 100)
                            {
                                tamplatePath += "Відомість обліку успішності (Державний екзамен) - протокол.xls";
                                DriverToDerzPas(tamplatePath);
                            }
                            else
                            {
                                tamplatePath += "Відомість обліку успішності (залік - диф. залік - екзамен).xls";
                                DriverToOblicOfZalic(tamplatePath);
                            }
                            _xlWorkBook.Close();
                        }
                    }
                    if (bl)
                    {
                        MessageBox.Show(string.Format("У {0}{1}", semestr, Resources.ExcelWork_CreateOblicUspishnosti_));
                        _xlApp.Quit();
                        return;
                    }
                    break;
                }

            }
            
            _xlApp.Quit();
        }

        private void ReloadSheet(string path)
        {
            var tamplateBook = _xlApp.Workbooks.Open(path);
            Excel.Worksheet tamplateSheet = (Excel.Worksheet)tamplateBook.Worksheets.Item[1];

            String nameSheet = _xlWorkSheet.Name;

            _xlApp.Visible = true;
            _xlApp.DisplayAlerts = false;

            //Переприсвоєння імені із видаленням листка
            tamplateSheet.Copy(_xlWorkSheet);
            
            _xlWorkSheet.Delete();

            _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[tamplateSheet.Name];
            _xlWorkSheet.Name = nameSheet;

            tamplateBook.Close();
            _xlWorkBook.Save();
        }

        private void DriverToDerzPas(String path)
        {
            ReloadSheet(path);

            _xlWorkSheet.Cells[4, "H"].Value = _subject;
            _xlWorkSheet.Cells[9, "C"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Код спеціальності"] + "_" + CurrentGroupName;
            for (int i = 0; i < DsRobPlan.Tables[CurrentGroupName].Rows.Count; i++)
            {
                if (DsRobPlan.Tables[CurrentGroupName].Rows[i][0].Equals(_subject))
                {
                    for (int j = 1; j < DsRobPlan.Tables[CurrentGroupName].Columns.Count; j++)
                    {
                        if (!DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("викладач")) continue;
                        //Викладач
                        _xlWorkSheet.Cells[20, "G"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j] + "_________________________________";
                        _xlWorkSheet.Cells[84, "H"].Value = "_____" + DsRobPlan.Tables[CurrentGroupName].Rows[i][j] + "__";
                    }
                    _xlWorkBook.Save();
                    break;
                }
            }
            int n = 46;
            foreach (DataRow dr in DsRobPlan.Tables["Студенти"].Rows)
            {
                if (dr[3].ToString().Equals(CurrentGroupName))
                {
                    _xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                    n++;
                }
            }

            //Кількість студентів у групі
            _xlWorkSheet.Cells[12, "G"] = "__" + (n - 46) + "__";

            _xlWorkBook.Save();
            if (n != 76)
                _xlWorkSheet.Range["B" + n, "Q" + 75].Delete();
            _xlWorkBook.Save();
        }

        private void DriverToKp(String path, bool practuca)
        {
            ReloadSheet(path);

            //Відділення
            var viddilenia = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Напрям підготовки"].ToString().Equals("Програмна інженерія") ? "Програмної інженерії" : "Метрології та інформаційно-вимірювальної технології";
            _xlWorkSheet.Cells[13, "E"].Value = viddilenia;

            //Спеціальність
            _xlWorkSheet.Cells[15, "F"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Спеціальність"].ToString();

            //Курс
            _xlWorkSheet.Cells[17, "D"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Курс"].ToString();

            //Група
            _xlWorkSheet.Cells[17, "G"].Value = CurrentGroupName;

            //Навчальний рік
            int year = Convert.ToInt32(ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"].ToString()) + 1;
            _xlWorkSheet.Cells[19, "I"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"] + "-" + year;

            //Назва дисципліни
            _xlWorkSheet.Cells[26, "F"].Value = _subject;

            //Семестр
            _xlWorkSheet.Cells[28, "D"].Value = _arabSemestr;

            //Номер відомості
            _xlWorkSheet.Cells[22, "M"].Value = _numberOfOblic;

            for (int i = 0; i < DsRobPlan.Tables[CurrentGroupName].Rows.Count; i++)
            {
                if (!DsRobPlan.Tables[CurrentGroupName].Rows[i][0].Equals(_subject)) continue;
                for (int j = 1; j < DsRobPlan.Tables[CurrentGroupName].Columns.Count; j++)
                {
                    if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("[семестр " + _arabSemestr + "]"))
                    {
                        if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("всього годин"))
                        {
                            //Всього годин
                            if (!DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString().Equals("0"))
                                _xlWorkSheet.Cells[30, "Q"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString();
                        }
                        else if (!DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString().Equals("0"))
                        {
                            //Форма здачі
                            if (!practuca) _xlWorkSheet.Cells[30, "F"].Value = "курсовий проект";
                            else
                            {
                                if (_subject.Trim().Equals("Технологічна практика") || _subject.Trim().Equals("Переддипломна практика"))
                                    _xlWorkSheet.Cells[30, "F"].Value = "захист";
                                else
                                    _xlWorkSheet.Cells[30, "F"].Value = "Д/З";
                            }
                        }
                    }
                    else if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("викладач"))
                    {
                        //Викладач
                        _xlWorkSheet.Cells[37, "K"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j] + "_____";
                        _xlWorkSheet.Cells[100, "N"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString();
                    }
                }
                _xlWorkBook.Save();
                break;
            }
            int n = 45;
            foreach (DataRow dr in DsRobPlan.Tables["Студенти"].Rows.Cast<DataRow>().Where(dr => dr[3].ToString().Equals(CurrentGroupName)))
            {
                _xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                _xlWorkSheet.Cells[n, "H"].Value = dr[2].ToString();
                n++;
            }
            _xlWorkBook.Save();
            if (n != 75)
                _xlWorkSheet.Range["B" + n, "Q" + 74].Delete();
            _xlWorkBook.Save();
        }

        private void DriverToOblicOfZalic(String path)
        {
            ReloadSheet(path);

            //Відділення
            string viddilenia;
            viddilenia = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Напрям підготовки"].ToString().Equals("Програмна інженерія") ? "Програмної інженерії" : "Метрології та інформаційно-вимірювальної технології";
            _xlWorkSheet.Cells[13, "E"].Value = viddilenia;

            //Спеціальність
            _xlWorkSheet.Cells[15, "F"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Спеціальність"].ToString();

            //Курс
            _xlWorkSheet.Cells[17, "D"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Курс"].ToString();

            //Група
            _xlWorkSheet.Cells[17, "G"].Value = CurrentGroupName;

            //Навчальний рік
            int year = Convert.ToInt32(ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"].ToString()) + 1;
            _xlWorkSheet.Cells[19, "I"].Value = ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"] + "-" + year;

            //Назва дисципліни
            _xlWorkSheet.Cells[26, "F"].Value = _subject;

            //Семестр
            _xlWorkSheet.Cells[28, "D"].Value = _arabSemestr;

            //Номер відомості
            _xlWorkSheet.Cells[22, "M"].Value = _numberOfOblic;

            for (int i = 0; i < DsRobPlan.Tables[CurrentGroupName].Rows.Count; i++)
            {
                if (DsRobPlan.Tables[CurrentGroupName].Rows[i][0].Equals(_subject))
                {
                    for (int j = 1; j < DsRobPlan.Tables[CurrentGroupName].Columns.Count; j++)
                    {
                        if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("[семестр " + _arabSemestr + "]"))
                        {
                            if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("всього годин"))
                            {
                                //Всього годин
                                _xlWorkSheet.Cells[30, "Q"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString();
                            }
                            else if (!DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString().Equals("0"))
                            {
                                //Форма здачі
                                _xlWorkSheet.Cells[30, "F"].Value =
                                    DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Substring(0, DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.IndexOf(" ", StringComparison.Ordinal));
                            }
                        }
                        else if (DsRobPlan.Tables[CurrentGroupName].Columns[j].ColumnName.Contains("викладач"))
                        {
                            //Викладач
                            _xlWorkSheet.Cells[32, "E"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString();
                            _xlWorkSheet.Cells[94, "N"].Value = DsRobPlan.Tables[CurrentGroupName].Rows[i][j].ToString();
                        }
                    }
                    _xlWorkBook.Save();
                    break;
                }
            }
            int n = 39;
            foreach (DataRow dr in DsRobPlan.Tables["Студенти"].Rows.Cast<DataRow>().Where(dr => dr[3].ToString().Equals(CurrentGroupName)))
            {
                _xlWorkSheet.Cells[n, "C"].Value = dr[1].ToString();
                _xlWorkSheet.Cells[n, "H"].Value = dr[2].ToString();
                n++;
            }
            _xlWorkBook.Save();
            if (n != 69)
            _xlWorkSheet.Range["B" + n, "Q" + 68].Delete();
            _xlWorkBook.Save();
        }

        private string ArabNormalize(string str)
        {
            char[] ch = str.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                var arg = (int)ch[i];
                if (arg == 1030) arg = 73;
                ch[i] = (char)arg;
            }
            return new string(ch);
        }

        private static string CutSheetName(string s, string arab)
        {
            var lenght = s.Length + arab.Length;
            if (lenght <= 32){
                return (s.Replace("*","&") + arab);
            }

            return (s.Substring(0, 31 - arab.Length).Replace("*", "&") + arab);
        }

        public DataTable DtSo = null;
        public void ZvedVidomist(int semestr, string subject, string mount)
        {
            _subject = subject;
            Semestr = semestr;


            _arabSemestr = semestr == 1 ? ArgDataSet.Tables[CurrentGroupName].Rows[0]["Перше півріччя"].ToString() : ArgDataSet.Tables[CurrentGroupName].Rows[0]["Друге півріччя"].ToString();

            string sT;
            if (mount.Equals(""))
            {
                sT = semestr == 1 ? "1-ше півріччя" : "2-ге півріччя";
            }
            else sT = mount + " місяць";

            var existsPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\" + "Empty book.xls";

            var pathTogroup = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Зведена відомість успішності\" + "Зведена відомість успішності за " + sT + ".xls";

            var pathToTamplateVidomist = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\" + "Зведена відомість.xls";

            if (!System.IO.File.Exists(pathTogroup))
                System.IO.File.Copy(existsPath, pathTogroup, true);

            _xlApp = new Excel.Application();
            _xlApp.DisplayAlerts = false;

            _xlWorkBook = _xlApp.Workbooks.Open(pathTogroup);
            _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[1];

            bool ifExist = true;
            for (int i = 1; i <= _xlWorkBook.Sheets.Count; i++)
            {
                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[i];
                if (_xlWorkSheet.Name.Contains("Sheet") || _xlWorkSheet.Name.Contains("Аркуш") || _xlWorkSheet.Name.Contains("Лист"))
                {
                    _xlWorkSheet.Name = CurrentGroupName;
                    ifExist = false;
                    _xlWorkBook.Save(); 
                    break;
                }
                if (!_xlWorkSheet.Name.Equals(CurrentGroupName)) continue;
                if (MessageBox.Show(Resources.ExcelWork_ZvedVidomist_Для_поточної_групи__ + CurrentGroupName + Resources.ExcelWork_ZvedVidomist_, "Обережно, ви можете втратити дані", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    _xlWorkBook.Close();
                    _xlApp.Quit();
                    return;
                }
                //xlWorkSheet.Delete();
                ifExist = false;
                //xlWorkBook.Save();
                break;
            }
            if (ifExist)
            {
                _xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _xlApp.DisplayAlerts = false;
                _xlWorkBook.Worksheets.get_Item(1).Name = CurrentGroupName;
                _xlWorkBook.Save();
            }
            _xlWorkSheet = _xlWorkBook.Worksheets.get_Item(CurrentGroupName);
            ReloadSheet(pathToTamplateVidomist);

            String[] forma = { "Іспит", "Д/З", "КП", "Залік", "КП" };
            char startX = 'E';
            char pos = ++startX;
            startX--;
            int count = 0;
            int otherCount = 0;
            for (int i = 0; i < forma.Length; i++)
            {
                bool bl = false;
                for (int row = 0; row < DsRobPlan.Tables[CurrentGroupName].Rows.Count; row++)
                {
                    try
                    {
                        if (!DsRobPlan.Tables[CurrentGroupName].Rows[row][forma[i] + " [семестр " + _arabSemestr + "]"].ToString().Equals("0") && i != 2 && i != 4)
                        {
                            bl = true;
                            startX++;
                            _xlWorkSheet.Cells[9, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                            _xlWorkSheet.Cells[43, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["викладач"].ToString();

                            _xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            _xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            _xlWorkSheet.Cells[9, startX.ToString()].ColumnWidth = ColumnWidth(DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());

                            count++;
                        }
                    }
                    catch (ArgumentException)
                    {
                        continue;
                    }
                    if (row == DsRobPlan.Tables[CurrentGroupName].Rows.Count - 1 && i != 2 && i != 4 && bl)
                    {
                        _xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Merge();
                        
                        _xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Value = forma[i];

                        _xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        _xlWorkSheet.Range[pos.ToString() + 8, startX.ToString() + 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                        pos = startX;
                        pos++;

                        bl = false;
                    }

                    if (!mount.Equals("")) continue;

                    if (!DsRobPlan.Tables[CurrentGroupName].Rows[row][forma[i] + " [семестр " + _arabSemestr + "]"].ToString().Equals("0") && i == 2 && i != 4
                        && !DsRobPlan.Tables[CurrentGroupName].Rows[row][forma[i] + " [семестр " + _arabSemestr + "]"].ToString().Equals("-1"))
                    {
                        startX++;
                        _xlWorkSheet.Cells[9, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                        _xlWorkSheet.Cells[43, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["викладач"].ToString();

                        _xlWorkSheet.Cells[9, startX.ToString()].HorizontalAlignment = Excel.XlLineStyle.xlContinuous;
                        _xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        _xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        _xlWorkSheet.Cells[43, startX.ToString()].ColumnWidth = ColumnWidth(DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());

                        count++;
                        pos = startX;
                        pos++;
                    }

                    if (i == 4 && DsRobPlan.Tables[CurrentGroupName].Rows[row][forma[i] + " [семестр " + _arabSemestr + "]"].ToString().Equals("-1"))
                    {
                        otherCount++;
                        startX++;
                        _xlWorkSheet.Cells[9, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString();
                        _xlWorkSheet.Cells[43, startX.ToString()].Value = DsRobPlan.Tables[CurrentGroupName].Rows[row]["викладач"].ToString();

                        _xlWorkSheet.Cells[9, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        _xlWorkSheet.Cells[43, startX.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        _xlWorkSheet.Cells[43, startX.ToString()].ColumnWidth = ColumnWidth(DsRobPlan.Tables[CurrentGroupName].Rows[row]["Назви навчальних  дисциплін"].ToString());
                        
                    }
                }
                _xlWorkBook.Save();
            }
            //добавлення та формування клітини з середнім балом
            startX++;
            _xlWorkSheet.Cells[9, startX.ToString()].Value = "Середній бал";
            _xlWorkSheet.Cells[9, startX.ToString()].ColumnWidth = 6.2;
            _xlWorkSheet.Cells[9, startX.ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            pos = 'E';
            for (int i = 0; i < count; i++)
                pos++;

            _xlWorkSheet.Range["F7", pos.ToString() + 7].Merge();
            _xlWorkSheet.Range["F7", pos.ToString() + 7].Value = "Предмети";
            _xlWorkSheet.Range["F7", pos.ToString() + 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            _xlWorkSheet.Range["F7", pos.ToString() + 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _xlWorkBook.Save();

            
            int startY = 11;
            int studCount = 0;
            foreach (DataRow dt in DsRobPlan.Tables["Студенти"].Rows.Cast<DataRow>().Where(dt => dt[3].ToString().Equals(CurrentGroupName)))
            {
                _xlWorkSheet.Range["D" + startY].Value = dt[1].ToString();
                if (dt[4].ToString().Equals("п")) _xlWorkSheet.Range["E" + startY].Value = dt[4].ToString();
                startY++;
                studCount++;
            }

            //куратор
            _xlWorkSheet.Range["K45"].Value = "/ " + CurrentCurator() + " /";

            if (startY != 40)
                _xlWorkSheet.Range["A" + startY, "IV" + 40].Delete();

            startY += 3;
            _xlWorkSheet.Range["C7", startX + startY.ToString()].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            if (mount.Equals(""))
            {
                _xlWorkSheet.Range["C4"].Value = "спеціальності \"" + ArgDataSet.Tables[CurrentGroupName].Rows[0]["Спеціальність"] + "\"";
                int year = Convert.ToInt32(ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"].ToString()) + 1;
                _xlWorkSheet.Range["D5"].Value = "групи " + CurrentGroupName + " за " + _arabSemestr + " семестр " +
                ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"] + "-" + year + " навчального року";
            }
            else
            {
                _xlWorkSheet.Range["D4"].Value = "спеціальності \"" + ArgDataSet.Tables[CurrentGroupName].Rows[0]["Спеціальність"] + "\"";
                int year = Convert.ToInt32(ArgDataSet.Tables[CurrentGroupName].Rows[0]["Рік"].ToString()) + 1;
                _xlWorkSheet.Range["C5"].Value = "за місяць " + mount + year + "р.";
                _xlWorkSheet.Range["C6"].Value = "група " + CurrentGroupName;
            }
            
            var pathToOblic  = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Облік успішності\" + CurrentGroupName + ".xls";

            _xlWorkBook.Save();
            //занесення оцінок
            if (System.IO.File.Exists(pathToOblic) && mount.Equals(""))
            {
                //прохід через усі записані у зведену відомість предмети
                DataTable dataTable = GetThePas(pathToOblic);

                startX = 'E';
                for (int i = 0; i < count + otherCount; i++)
                {
                    startX++;
                    string likeSheet = CutSheetName(_xlWorkSheet.Range[startX.ToString() + 9].Value.ToString(), "_" + _arabSemestr + "_");
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
                        catch (Exception)
                        {
                            break;
                        }
                        currentRow++;
                        var pas = dataTable.Rows[currentRow - 1][likeSheet].ToString();
                        int cell = j + 10;
                        _xlWorkSheet.Range[startX.ToString() + cell].Value = pas;
                        if (pas == null || pas.Equals("") || pas.Equals(" "))
                        {
                            continue;
                        }

                        if (pas.Contains("з")) continue;
                        try
                        {
                            int number = Convert.ToInt32(pas);
                            if (number <= 3) superNegativPasCount++;
                            else if (number < 7) negatPasCount++;
                        }
                        catch (Exception)
                        {
                            // ignored
                        }
                    }
                    _xlWorkSheet.Range[startX.ToString() + (studCount + 11)].Formula = "=" + (studCount - superNegativPasCount) + "/" + studCount;
                    _xlWorkSheet.Range[startX.ToString() + (studCount + 12)].Formula = "=" + (studCount - (superNegativPasCount + negatPasCount)) + "/" + studCount;
                }
                _xlWorkBook.Save();
                startX = 'E';

                for (int i = 0; i < count + otherCount; i++)
                    startX++;

                char begin = startX;
                char averageBal = ++startX;
                startX++;
                
                for (int i = 11; i < studCount + 11; i++)
                {
                    _xlWorkSheet.Range[averageBal.ToString() + i].Formula = "=AVERAGE(" + "F" + i + ":" + begin.ToString() + i + ") - 0.5";
                    _xlWorkSheet.Range[averageBal.ToString() + i].NumberFormatLocal = "##";
                    string s1 = _xlWorkSheet.Range["E" + i].Value;
                    double s2;
                    try
                    {
                        s2 = _xlWorkSheet.Range[averageBal.ToString() + i].Value;
                    }
                    catch (Exception) { continue; }

                    if (s1 == null)
                    {
                        if (s2 > 7)
                        {
                            _xlWorkSheet.Range[startX.ToString() + i].Value = "1";
                            if (dataTable.Rows[i - 11]["підвищена стипендія"].ToString().Equals("1"))
                                _xlWorkSheet.Range[startX.ToString() + i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }
                }
                _xlWorkBook.Save();
            }

            startX = 'E';

            for (int i = 0; i < count + otherCount + 1; i++)
                startX++;

            char stypendia = ++startX;
            startX++;
            int xlIter = 10;
            foreach (DataRow dt in DsRobPlan.Tables["Студенти"].Rows.Cast<DataRow>().Where(dt => dt[3].ToString().Equals(CurrentGroupName)))
            {
                xlIter++;
                var whatType = dt[0];
                if (whatType != null)
                {
                    bool bl = true;
                    var whatype2 = _xlWorkSheet.Range[stypendia.ToString() + xlIter].Value;
                    if (whatype2 == null) bl = false;

                    if (whatType.ToString().Contains("сир"))
                        _xlWorkSheet.Range[startX.ToString() + xlIter].Value = "с";
                    else if (whatType.ToString().Contains("гір") && bl)
                        _xlWorkSheet.Range[startX.ToString() + xlIter].Value = "г";
                    else if (whatType.ToString().Contains("інва"))
                        _xlWorkSheet.Range[startX.ToString() + xlIter].Value = "і";
                }
            }
            _xlWorkBook.Save();
            _xlWorkBook.Close();
            _xlApp.Quit();
        }

        private DataTable GetThePas(String path)
        {
            DataTable dt = new DataTable();

            var newBook = _xlApp.Workbooks.Open(path);

            dt.Columns.Add("ID", typeof(Int32));
            dt.Columns[0].AllowDBNull = false;
            dt.Columns[0].AutoIncrement = true;
            dt.Columns[0].AutoIncrementStep = 1;
            dt.Columns[0].Unique = true;

            for (int i = 1; i <= newBook.Worksheets.Count; i++)
            {
                dt.Columns.Add(newBook.Worksheets.get_Item(i).Name.ToString(), typeof(string));
            }
            
            String [] st = {"B", "J", "46", "B", "L", "39", "B", "L", "45"};
            int pos = 0;
            for (int i = 1; i <= newBook.Worksheets.Count; i++)
            {
                var newSheet = (Excel.Worksheet)newBook.Worksheets.Item[i];
                for (int j = 0; j < st.Length; j += 3)
                {
                    var whatType = newSheet.Range[st[j] + st[j + 2]].Value;
                    if (whatType == null) continue;
                    if (whatType.Equals("1."))
                    { 
                        pos = j;
                        break;
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
                    var ocinka = newSheet.Range[st[pos + 1] + j].Value ?? " ";
                    dt.Rows[j - Convert.ToInt32(st[pos + 2])][newSheet.Name] = ocinka.ToString();
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
                    catch (Exception)
                    {
                        // ignored
                    }
                }

                if (ifHight)
                    dt.Rows[i]["підвищена стипендія"] = "1";
            }

            newBook.Close();
            return dt;
        }

        private static double ColumnWidth(string s)
        {
            if (s.Length <= 21) return 5.57;
            if (s.Length <= 40) return 9.70;
            if (s.Length <= 55) return 11;
            return 13.43;
        }

        public void ArhiveZvedVid(string path)
        {
            if (!path.Contains("Зведена відомість успішності за"))
            {
                MessageBox.Show(Resources.ExcelWork_ArhiveZvedVid_);
                return;
            }

            _xlApp = new Excel.Application();
            _xlApp.DisplayAlerts = false;
            _xlWorkBook = _xlApp.Workbooks.Open(path);
            
            String groupArxivePath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"User Data\Зведена відомість успішності\Архів\";

            String emptyBookPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                + @"Data\Empty book.xls";

            for (int i = 1; i <= _xlWorkBook.Worksheets.Count; i++)
            {
                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[i];
                String sheetName = _xlWorkSheet.Name;

                if (!System.IO.File.Exists(groupArxivePath + sheetName + ".xls"))
                {
                    int semestCount = 6;
                    if (sheetName.Contains("ПІ")) semestCount = 8;

                    System.IO.File.Copy(emptyBookPath, groupArxivePath + sheetName + ".xls", true);

                    Excel.Workbook bookNew = _xlApp.Workbooks.Open(groupArxivePath + sheetName + ".xls");

                    for (int j = 1; j <= semestCount; j++)
                    {
                        bookNew.Worksheets.Add();
                        bookNew.Save();
                        bookNew.Worksheets.get_Item(1).Name = "семестр " + ArabToRome(j);
                        bookNew.Save();
                    }
                    bookNew.Close();

                }
                ReloadSheet2(groupArxivePath + sheetName + ".xls");
            }
        }

        private void ReloadSheet2(String path)
        {
            Excel.Workbook bookNew = _xlApp.Workbooks.Open(path);

            for (int j = 1; j <= bookNew.Worksheets.Count - 1; j++)
            {
                Excel.Worksheet sheetNew = (Excel.Worksheet)bookNew.Worksheets.Item[j];

                String semestr = _xlWorkSheet.Range["D5"].Value.ToString().Substring(18);
                semestr = semestr.Substring(0, semestr.IndexOf(" ", StringComparison.Ordinal));
                //MessageBox.Show(">" + semestr + "<\n");

                String sheet = sheetNew.Name.Substring(sheetNew.Name.IndexOf(" ", StringComparison.Ordinal) + 1);
                //semestr.Substring(0, semestr.IndexOf(" "));

                //MessageBox.Show(">" + sheet + "<\n" + ">" + semestr + "<");
                if (semestr.Equals(sheet))
                {
                    //Переприсвоєння імені із видаленням листка
                    _xlWorkSheet.Copy(sheetNew);
                    bookNew.Save();
                    String name = sheetNew.Name;

                    sheetNew.Application.DisplayAlerts = false;
                    sheetNew.Delete();
                    bookNew.Save();

                    try
                    {
                        sheetNew = (Excel.Worksheet)bookNew.Worksheets.Item[name];
                    }
                    catch (Exception ex)
                    {
                        sheetNew = (Excel.Worksheet)bookNew.Worksheets.Item[_xlWorkSheet.Name];
                        sheetNew.Name = name;

                        bookNew.Save();
                    }
                }
            }
            bookNew.Close();
        }

        public String ArabToRome(int arab)
        {
            var rome = "";
            switch (arab)
            {
                case 1:
                    rome = "I";
                    break;
                case 2:
                    rome = "II";
                    break;
                case 3:
                    rome = "III";
                    break;
                case 4:
                    rome = "IV";
                    break;
                case 5:
                    rome = "V";
                    break;
                case 6:
                    rome = "VI";
                    break;
                case 7:
                    rome = "VII";
                    break;
                case 8:
                    rome = "VIII";
                    break;
            }
            return rome;
        }

        public string CurrentCurator()
        {
            String path = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.Length - 9)
                                 + @"Data\Куратори.xls";

            var book = _xlApp.Workbooks.Open(path);
            var sheet = (Excel.Worksheet)book.Worksheets.Item["Куратори"];

            for (int i = 2; ; i++)
            {
                if (sheet.Range["A" + i].Value == null) break;
                if (sheet.Range["A" + i].Value.ToString().Equals(CurrentGroupName))
                {
                    return sheet.Range["B" + i].Value.ToString();
                }
            }

            book.Close();
            return "";
        }
    }
}