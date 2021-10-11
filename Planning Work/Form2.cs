using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Security.Policy;
using System.Runtime.InteropServices;
using System.Data.SqlClient;


namespace Planning_Work
{
    public partial class Form2 : Form
    {
        public int cutCount, cutRow,cutColumn,row, colmn, countFile, allRow, allColumn, ok = 0, countFirst, countTwo, countTree, countFour, countFive, indexRow, indexColumn, ItIsI = 0, hours, hour, DragRow, DragColumn, countGropeForSql;
        public string numberGroup, disciplines, tema, room, numberGroupCut;
        public string [] teach;
        public bool[] numberWork = new bool[6];
        public bool _pasteOrCut = false;
        public AllLessinAndRooms clas;
        public Teacher teacher;
        public AllLessin[] lessons;
        public DataGridViewComboBoxCell combo_cell = new DataGridViewComboBoxCell();
        public DateTime time;
        public Allpeople people;
        public CellsTable[,] arrayTable;
        public CellsTable cut_Buufer;
        public AllColorDisipilens arrayColorDisiplines;
        public TeacherForSql sqlSupTeacher = new TeacherForSql();

        public Form2(int countUploadFile, DateTime times)
        {
            if (countUploadFile >= 0)
            {

                GetOtherForSQL();
                lessons = GetLessingForSQL();
                clas = GetAllRoomsForSQL();
                people = GetPeopleForSQL();
                teacher = GetTeacherForSQL();
                clas.countGroupe();


                InitializeComponent();
                label2.MaximumSize = new Size(200,800);
                KeyPreview = true;
                WindowState = FormWindowState.Maximized;
                colorDialog1.Color = Color.Red;
                colorDialog2.Color = Color.Red;
                colorDialog3.Color = Color.Red;
                colorDialog4.Color = Color.Red;
                indexColumn = 0;
                indexRow = 0;
                time = times;
                countFile = countUploadFile;

                countFirst = people.BigBoss() * 2;
                countTwo = people.BigBoss() * 2;
                countTree = people.BigBoss() * 2;
                countFour = people.BigBoss() * 2;
                countFive = people.BigBoss() * 2;

                arrayTable = new CellsTable[allRow + 1, allColumn + 1];
                for (int i = 0; i < allColumn; i++)
                {
                    dataGridView1.Columns.Add(Convert.ToString(i), "");
                }

                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                for (int j = 0; j < allRow; j++)
                {
                    dataGridView1.Rows.Add("");
                }

                createText(dataGridView1, people);
                setClassAllLessin(dataGridView1, clas);

                dataGridView1.Rows[0].ReadOnly = true;
                dataGridView1.Rows[1].ReadOnly = true;
                dataGridView1.Rows[2].ReadOnly = true;
                dataGridView1.Rows[3].ReadOnly = true;
                dataGridView1.Rows[4].ReadOnly = true;

                dataGridView1.Columns[0].ReadOnly = true;
                dataGridView1.Columns[1].ReadOnly = true;

                dataGridView1.Columns[2 + countFirst].ReadOnly = true;
                dataGridView1.Columns[3 + countFirst].ReadOnly = true;

                dataGridView1.Columns[4 + countFirst + countTwo].ReadOnly = true;
                dataGridView1.Columns[5 + countFirst + countTwo].ReadOnly = true;

                dataGridView1.Columns[6 + countFirst + countTwo + countTree].ReadOnly = true;
                dataGridView1.Columns[7 + countFirst + countTwo + countTree].ReadOnly = true;

                dataGridView1.Columns[8 + countFirst + countTwo + countTree + countFour].ReadOnly = true;
                dataGridView1.Columns[9 + countFirst + countTwo + countTree + countFour].ReadOnly = true;


                arrayColorDisiplines = new AllColorDisipilens(GetCountUniversalDisiplines(lessons));
                dataGridView1.AllowDrop = true;
                ColorDataGrid();
                VOSKRESENIE();
                getAlllogical();//Запись буферных зон (Что бы ты еблан не искал потом 1 милион лет)
            }
            else
            {
                MessageBox.Show("Загружено " + Convert.ToString(countUploadFile) + ". \nЗагрузите все 5 курсов для корректной работы програмы.");
                InitializeComponent();
                Close();
            }


        }

        //Подключение к таблице 
        public string get_cs()
        {
            return "Data Source=" + Properties.Settings.Default.PathSqlServer + "; Initial Catalog =DarkLight ; User ID = sa; Password = 123456";
        }

        //Получение row, column и countGroupe
        private void GetOtherForSQL()
        {
            string sqlExpression = "SELECT * FROM _other";
            using (var connection = new SqlConnection(get_cs()))
            {
                connection.Open();
                {
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader read = command.ExecuteReader();

                    if (read.HasRows)
                    {
                        while (read.Read())
                        {
                            allRow = Convert.ToInt32(read.GetValue(1).ToString());
                            allColumn = Convert.ToInt32(read.GetValue(2).ToString());
                            countGropeForSql = Convert.ToInt32(read.GetValue(3).ToString());
                        }
                    }
                }
                connection.Close();
            }
        }

        //Получение AllLesing
        private AllLessin[] GetLessingForSQL()
        {

            bool start = true;
            int i = 0;
            string sqlExpression = "SELECT * FROM _allPeolpe";
            AllLessin[] answer = new AllLessin[countGropeForSql];
            using (var connection = new SqlConnection(get_cs()))
            {
                connection.Open();
                {
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader read = command.ExecuteReader();

                    if (read.HasRows)
                    {
                        while (read.Read())
                        {
                            if (start)
                            {
                                Triple push;
                                push._fack = Convert.ToInt32(read.GetValue(2).ToString());
                                push._nameDiscipline = read.GetValue(3).ToString();
                                push._comentsDisciplines = read.GetValue(4).ToString(); ;
                                push._tema = read.GetValue(5).ToString();
                                push._coments = read.GetValue(6).ToString();
                                push._time = read.GetValue(7).ToString();
                                push._set = Convert.ToBoolean(read.GetValue(8).ToString());
                                answer[i] = new AllLessin(read.GetValue(1).ToString(), 0);
                                answer[i].pushOneTriple(push);
                                start = false;
                            }
                            else
                            {
                                Triple push;
                                push._fack = Convert.ToInt32(read.GetValue(2).ToString());
                                push._nameDiscipline = read.GetValue(3).ToString();
                                push._comentsDisciplines = read.GetValue(4).ToString(); ;
                                push._tema = read.GetValue(5).ToString();
                                push._coments = read.GetValue(6).ToString();
                                push._time = read.GetValue(7).ToString();
                                push._set = Convert.ToBoolean(read.GetValue(8).ToString());
                                if (read.GetValue(1).ToString() == answer[i]._nameAllLessin)
                                {
                                    answer[i].pushOneTriple(push);
                                }
                                else
                                {
                                    i++;
                                    answer[i] = new AllLessin(read.GetValue(1).ToString(), 0);
                                    answer[i].pushOneTriple(push);
                                }
                            }

                        }
                    }
                }
                connection.Close();
            }


            return answer;
        }

        //Получение Allpeople
        private Allpeople GetPeopleForSQL()
        {
            string sqlExpression = "SELECT * FROM _baseItems";
            Allpeople answer = new Allpeople();

            using (var connection = new SqlConnection(get_cs()))
            {
                connection.Open();
                {
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader read = command.ExecuteReader();

                    if (read.HasRows)
                    {
                        while (read.Read())
                        {
                            answer.pushElements(Convert.ToString(read.GetValue(3)), Convert.ToInt32(read.GetValue(1)), Convert.ToString(read.GetValue(4)));
                        }
                    }
                }
                connection.Close();
            }
            return answer;
        }

        //Получение class
        private AllLessinAndRooms GetAllRoomsForSQL()
        {
            string sqlExpression = "SELECT * FROM _class";
            AllLessinAndRooms answer = new AllLessinAndRooms();
            using (var connection = new SqlConnection(get_cs()))
            {
                connection.Open();
                {
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader read = command.ExecuteReader();

                    if (read.HasRows)
                    {
                        while (read.Read())
                        {
                            answer.pushGPSQL(Convert.ToString(read.GetValue(2)), Convert.ToString(read.GetValue(1)), Convert.ToString(read.GetValue(3)));
                        }
                    }
                }
                connection.Close();
            }

            return answer;
        }

        //Получение teacher
        private Teacher GetTeacherForSQL()
        {
            Teacher answer = new Teacher();
            string sqlExpression = "SELECT * FROM _teacher";
            using (var connection = new SqlConnection(get_cs()))
            {
                connection.Open();
                {
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader read = command.ExecuteReader();

                    if (read.HasRows)
                    {
                        while (read.Read())
                        {
                            sqlSupTeacher.push(Convert.ToString(read.GetValue(1)), Convert.ToString(read.GetValue(2)), Convert.ToString(read.GetValue(3)), Convert.ToString(read.GetValue(4)));
                        }
                    }
                }
                connection.Close();
            }


            for (int i = 0; i < sqlSupTeacher._array.Length;)
            {
                int countBuffer = countOneGroupe_SQL(sqlSupTeacher, i);
                string[] buffer = new string[countBuffer];
                string nameGroupeBuffer = sqlSupTeacher._array[i]._nameGroup;
                string nameDisiplinesBuffer = sqlSupTeacher._array[i]._nameDis;
                string teamBuffer = sqlSupTeacher._array[i]._team;
                for (int j = 0; j < countBuffer; j++)
                {
                    nameGroupeBuffer = sqlSupTeacher._array[i]._nameGroup;
                    nameDisiplinesBuffer = sqlSupTeacher._array[i]._nameDis;
                    buffer[j] = sqlSupTeacher._array[i]._teacher;
                    if (i < sqlSupTeacher._array.Length)
                    {
                        i++;
                    }
                }
                answer.push(nameGroupeBuffer, nameDisiplinesBuffer, teamBuffer, buffer);
            }

            return answer;
        }

        //Для Sql
        public int countOneGroupe_SQL(TeacherForSql teach, int number)
        {
            int count = 1;
            for (int i = number; i < teach._array.Length; i++)
            {
                if (i + 1 < teach._array.Length && teach._array[i]._nameDis == teach._array[i + 1]._nameDis && teach._array[i]._nameGroup == teach._array[i + 1]._nameGroup)
                {
                    count++;
                }
                else
                {
                    return count;
                }
            }

            return count;
        }

        //Кнопка форматировании ячейки
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.TopCenter;
            dataGridView1.AutoResizeColumns();
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        }

        // Русификотор дней недели
        public string RussFuckingDays(string days)
        {
            string RusDAY = "";

            if (days == "Monday")
                RusDAY = "пн";
            if (days == "Tuesday")
                RusDAY = "вт";
            if (days == "Wednesday")
                RusDAY = "ср";
            if (days == "Thursday")
                RusDAY = "чт";
            if (days == "Friday")
                RusDAY = "пт";
            if (days == "Saturday")
                RusDAY = "сб";
            if (days == "Sunday")
                RusDAY = "вс";


            return RusDAY;
        }

        //Форматированеие всей страницы заполненеия
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            createDate(e, people, countFile, 0, 0);
        }

        //Сохранение (вывод все в таблицу)
        private void button4_Click(object sender, EventArgs e)
        {
            string pathFile = "";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathFile = saveFileDialog1.FileName;


                Excel.Application App;
                Excel.Workbook xlsWB;
                Excel.Worksheet xlsSheet;

                App = new Excel.Application();
                xlsWB = App.Workbooks.Add();
                xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);

                App = new Excel.Application();

                for (int i = 0; i < allRow; i++)
                {
                    for (int j = 0; j < allColumn; j++)
                    {
                        if (dataGridView1[j, i].Value != null)
                        {
                            xlsSheet.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                        }
                    }
                }
                xlsWB.SaveAs(pathFile);
                xlsWB.Close(true);
                App.Quit();
            }
        }

        //Размер ячеек и текста
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1[colmn, row].ReadOnly != true && dataGridView1[colmn, row].Value.ToString() != "")
            {
                dataGridView1[colmn, row].Value = "";
                arrayTable[row, colmn]._disiplines = null;
                arrayTable[row, colmn]._rooms = null;
                arrayTable[row, colmn]._tema = null;
                arrayTable[row, colmn]._teacher = null;
                cheakGorizontal(row);
                cheakVertical(colmn);


                for (int i = 0; i < allRow; i++)
                {
                    for (int j = 0; j < allColumn; j++)
                    {
                        if (arrayTable[i, j]._disiplines != null)
                        {
                            dataGridView1[j, i].Style.ForeColor = Color.Black;
                            dataGridView1[j, i].ToolTipText = "";
                            cheakVertical(j);
                            cheakGorizontal(i);
                        }
                    }
                }
            }
        }

        // Обеденяет ячейки сверху вних 
        public void AllLessinCellsRow(DataGridViewCellPaintingEventArgs e, int column, int start_row, int back_row)
        {
            if (e.RowIndex >= start_row && e.ColumnIndex == column && e.RowIndex < back_row)
            {
                e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        // Обеденяет ячейки справа на лево 
        public void AllLessinCellsColumn(DataGridViewCellPaintingEventArgs e, int row, int start_column, int back_column)
        {
            if (e.RowIndex == row && e.ColumnIndex >= start_column && e.ColumnIndex <= back_column)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        // Рисует ячейку одного дня
        public void createOneDaysFormFactors(DataGridViewCellPaintingEventArgs e, int row, int column)
        {
            AllLessinCellsRow(e, column, row, row + 3);
        }

        //Цвет совпадения преподавателя
        private void button5_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();

        }

        //Цвет совпадения аудиторий
        private void button6_Click(object sender, EventArgs e)
        {
            colorDialog2.ShowDialog();
        }

        //Цвет данному предмету
        private void button7_Click(object sender, EventArgs e)
        {
            //if (checkBox1.Checked)
            //{
            //    MessageBox.Show("Отключите режим перетаскивания");
            //}
            //else
            //{
                colorDialog3.ShowDialog();
                arrayColorDisiplines.PushColorDisiplines(colorDialog3.Color, arrayTable[row, colmn]._disiplines);
                for (int i = 0; i < allRow; i++)
                {
                    for (int j = 0; j < allColumn; j++)
                    {
                        if (dataGridView1[j, i].Value != null && arrayTable[i, j]._disiplines == arrayTable[row, colmn]._disiplines)
                        {
                            dataGridView1[j, i].Style.BackColor = colorDialog3.Color;
                        }
                    }
                }
            //}

        }

        //Цвет нарушения логики 
        private void button8_Click(object sender, EventArgs e)
        {
            colorDialog4.ShowDialog();
        }

        //Создание оболочки формы
        public void createDate(DataGridViewCellPaintingEventArgs e, Allpeople people, int countUploadFile, int row, int column)
        {
            if (e.RowIndex >= 5)
            {
                row = 5;
                column = 0;

                for (; row <= allRow; row += 4)
                {
                    createOneDaysFormFactors(e, row, column);

                    createOneDaysFormFactors(e, row, column + countFirst + 2);
                    createOneDaysFormFactors(e, row, column + countFirst + countTwo + 4);
                    createOneDaysFormFactors(e, row, column + countFirst + countTwo + countTree + 6);
                    createOneDaysFormFactors(e, row, column + countFirst + countTwo + countTree + countFour + 8);
                }
            }
            if (e.RowIndex <= 5)
            {
                int start_row = row;
                int start_column = column;
                row = start_row;
                column = start_column;

                AllLessinCellsRow(e, column, row, row + 3);
                AllLessinCellsRow(e, column + countFirst + 2, row, row + 3);
                AllLessinCellsRow(e, column + countFirst + countTwo + 4, row, row + 3);
                AllLessinCellsRow(e, column + countFirst + countTwo + countTree + 6, row, row + 3);
                AllLessinCellsRow(e, column + countFirst + countTwo + countTree + countFour + 8, row, row + 3);

                AllLessinCellsRow(e, column + 1, row + 1, row + 4);
                AllLessinCellsRow(e, column + countFirst + 3, row + 1, row + 4);
                AllLessinCellsRow(e, column + countFirst + countTwo + 5, row + 1, row + 4);
                AllLessinCellsRow(e, column + countFirst + countTwo + countTree + 7, row + 1, row + 4);
                AllLessinCellsRow(e, column + countFirst + countTwo + countTree + countFour + 9, row + 1, row + 4);



                if (e.ColumnIndex < 2 + countFirst)
                {
                    AllLessinCellsColumn(e, 0, 2, countFirst);
                }


                if (e.ColumnIndex < 4 + countTwo + countFirst - 1 && e.ColumnIndex > 2 + countFirst)
                {
                    AllLessinCellsColumn(e, 0, 4 + countFirst, 4 + countFirst + countTwo);
                }

                if (e.ColumnIndex < 6 + countTree + countTwo + countFirst - 1 && e.ColumnIndex > 4 + countFirst + countTwo)
                {
                    AllLessinCellsColumn(e, 0, 6 + countFirst + countTwo, 6 + countFirst + countTwo + countTree);
                }

                if (e.ColumnIndex < 8 + countFour + countTree + countTwo + countFirst - 1 && e.ColumnIndex > 6 + countFirst + countTwo + countTree)
                {
                    AllLessinCellsColumn(e, 0, 8 + countFirst + countTwo + countTree, 8 + countFirst + countTwo + countTree + countFour);
                }

                if (e.ColumnIndex < 10 + countFive + countFour + countTree + countTwo + countFirst - 1 && e.ColumnIndex > 8 + countFirst + countTwo + countTree + countFour)
                {
                    AllLessinCellsColumn(e, 0, 10 + countFirst + countTwo + countTree + countFour, 10 + countFirst + countTwo + countTree + countFour + countFive);
                }
            }
        }

        // Cоздание текстовой оболочки
        public void createText(DataGridView dataGridView1, Allpeople people)
        {

            string dateOne;
            int row = 6;
            int column = 0;
            for (int i = 0; i < DateTime.DaysInMonth(time.Year, time.Month); i++)
            {
                dateOne = "";
                dateOne = dateOne + (i + 1).ToString() + "." + time.Month.ToString() + "." + time.Year.ToString();

                
                dataGridView1[column, row + i * 4].Value = dateOne;
                dataGridView1[column + 2 + countFirst, row + i * 4].Value = dateOne;
                dataGridView1[column + 4 + countFirst + countTwo, row + i * 4].Value = dateOne;
                dataGridView1[column + 6 + countFirst + countTwo + countTree, row + i * 4].Value = dateOne;
                dataGridView1[column + 8 + countFirst + countTwo + countTree + countFour, row + i * 4].Value = dateOne;
                DateTime buff = DateTime.Parse(dateOne);

                dataGridView1[column, row + i * 4 + 1].Value = RussFuckingDays(buff.DayOfWeek.ToString());
                dataGridView1[column + countFirst + 2, row + i * 4 + 1].Value = RussFuckingDays(buff.DayOfWeek.ToString());
                dataGridView1[column + countFirst + 4 + countTwo, row + i * 4 + 1].Value = RussFuckingDays(buff.DayOfWeek.ToString());
                dataGridView1[column + countFirst + 6 + countTwo + countTree, row + i * 4 + 1].Value = RussFuckingDays(buff.DayOfWeek.ToString());
                dataGridView1[column + countFirst + 8 + countTwo + countTree + countFour, row + i * 4 + 1].Value = RussFuckingDays(buff.DayOfWeek.ToString());

                dataGridView1[column + 1, row + i * 4 - 1].Value = "1 \n 2";
                dataGridView1[column + 3 + countFirst, row + i * 4 - 1].Value = "1 \n 2";
                dataGridView1[column + 5 + countFirst + countTwo, row + i * 4 - 1].Value = "1 \n 2";
                dataGridView1[column + 7 + countFirst + countTwo + countTree, row + i * 4 - 1].Value = "1 \n 2";
                dataGridView1[column + 9 + countFirst + countTwo + countTree + countFour, row + i * 4 - 1].Value = "1 \n 2";

                dataGridView1[column + 1, row + i * 4].Value = "3 \n 4";
                dataGridView1[column + 3 + countFirst, row + i * 4].Value = "3 \n 4";
                dataGridView1[column + 5 + countFirst + countTwo, row + i * 4].Value = "3 \n 4";
                dataGridView1[column + 7 + countFirst + countTwo + countTree, row + i * 4].Value = "3 \n 4";
                dataGridView1[column + 9 + countFirst + countTwo + countTree + countFour, row + i * 4].Value = "3 \n 4";

                dataGridView1[column + 1, row + i * 4 + 1].Value = "5 \n 6";
                dataGridView1[column + 3 + countFirst, row + i * 4 + 1].Value = "5 \n 6";
                dataGridView1[column + 5 + countFirst + countTwo, row + i * 4 + 1].Value = "5 \n 6";
                dataGridView1[column + 7 + countFirst + countTwo + countTree, row + i * 4 + 1].Value = "5 \n 6";
                dataGridView1[column + 9 + countFirst + countTwo + countTree + countFour, row + i * 4 + 1].Value = "5 \n 6";

                dataGridView1[column + 1, row + i * 4 + 2].Value = "7 \n 8";
                dataGridView1[column + 3 + countFirst, row + i * 4 + 2].Value = "7 \n 8";
                dataGridView1[column + 5 + countFirst + countTwo, row + i * 4 + 2].Value = "7 \n 8";
                dataGridView1[column + 7 + countFirst + countTwo + countTree, row + i * 4 + 2].Value = "7 \n 8";
                dataGridView1[column + 9 + countFirst + countTwo + countTree + countFour, row + i * 4 + 2].Value = "7 \n 8";
            }
            column = 0;
            row = 0;

            dataGridView1[column + 1, row + 2].Value = "Ч";
            dataGridView1[column, row + 1].Value = "Дата";
            dataGridView1[column, row + 4].Value = "Уч класс";
            dataGridView1[column, row + 1].Value = "Дата";

            dataGridView1[column + 3 + countFirst, row + 2].Value = "Ч";
            dataGridView1[column + 2 + countFirst, row + 1].Value = "Дата";
            dataGridView1[column + 2 + countFirst, row + 4].Value = "Уч класс";
            dataGridView1[column + 2 + countFirst, row + 1].Value = "Дата";

            dataGridView1[column + 5 + countFirst + countTwo, row + 2].Value = "Ч";
            dataGridView1[column + 4 + countFirst + countTwo, row + 1].Value = "Дата";
            dataGridView1[column + 4 + countFirst + countTwo, row + 4].Value = "Уч класс";
            dataGridView1[column + 4 + countFirst + countTwo, row + 1].Value = "Дата";

            dataGridView1[column + 7 + countFirst + countTwo + countTree, row + 2].Value = "Ч";
            dataGridView1[column + 6 + countFirst + countTwo + countTree, row + 1].Value = "Дата";
            dataGridView1[column + 6 + countFirst + countTwo + countTree, row + 4].Value = "Уч класс";
            dataGridView1[column + 6 + countFirst + countTwo + countTree, row + 1].Value = "Дата";

            dataGridView1[column + 9 + countFirst + countTwo + countTree + countFour, row + 2].Value = "Ч";
            dataGridView1[column + 8 + countFirst + countTwo + countTree + countFour, row + 1].Value = "Дата";
            dataGridView1[column + 8 + countFirst + countTwo + countTree + countFour, row + 4].Value = "Уч класс";
            dataGridView1[column + 8 + countFirst + countTwo + countTree + countFour, row + 1].Value = "Дата";

            people.GetOneAllLessin(dataGridView1, 3, 2, countFirst / 2, allRow, arrayTable);
            people.GetOneAllLessin(dataGridView1, 3, 4 + countFirst, countTwo / 2, allRow, arrayTable);
            people.GetOneAllLessin(dataGridView1, 3, 6 + countFirst + countTwo, countTree / 2, allRow, arrayTable);
            people.GetOneAllLessin(dataGridView1, 3, 8 + countFirst + countTwo + countTree, countFour / 2, allRow, arrayTable);
            people.GetOneAllLessin(dataGridView1, 3, 10 + countFirst + countTwo + countTree + countFour, countFive / 2, allRow, arrayTable);

            dataGridView1[2, 0].Value = "1 Курс";
            dataGridView1[4 + countFirst, 0].Value = "2 Курс";
            dataGridView1[6 + countFirst + countTwo, 0].Value = "3 Курс";
            dataGridView1[8 + countFirst + countTwo + countTree, 0].Value = "4 Курс";
            dataGridView1[10 + countFirst + countTwo + countTree + countFour, 0].Value = "5 Курс";
        }
        
        //Закрытие
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        // Заполняет таблицу (непонятно чем)
        public void setClassAllLessin(DataGridView e, AllLessinAndRooms clas)
        {
            foreach (AllLessinRomms elements in clas._array)
            {
                int i = 0;
                while (i < e.ColumnCount)
                {
                    if (e[i, 3].Value != null && e[i, 3].Value.ToString() == elements.AllLessin)
                    {
                        e[i, 4].Value = elements.rooms;
                    }
                    i++;
                }
            }
        }

        // Drag and Drop Услровия
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            int SourseRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
            int SourseColumn = dataGridView1.HitTest(e.X, e.Y).ColumnIndex;
            if (SourseRow > -1 && SourseColumn  > -1)
            {
                row = SourseRow;
                colmn = SourseColumn;
                if (!checkBox1.Checked == true && !dataGridView1[colmn, row].ReadOnly)
                {

                    DragRow = SourseRow;
                    DragColumn = SourseColumn;
                    dataGridView1.DoDragDrop(DragRow, DragDropEffects.Copy);
                }
            }
            
        }

        // Drag abd drop эффект
        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.C && e.Control) && !checkBox1.Checked)
            {
                button9.PerformClick();
            }

            if (e.KeyCode == Keys.X && e.Control)
            {
                if (checkBox1.Checked)
                {
                    checkBox1.Checked = false;
                }
                else
                {
                    checkBox1.Checked = true;

                }
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            button9.PerformClick();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //Вырезка
        private void button9_Click(object sender, EventArgs e)
        {
            if (!dataGridView1[colmn, row].ReadOnly)
            {
                if (!_pasteOrCut)
                {
                    if (arrayTable[row, colmn]._teacher != null)
                    {
                        if (dataGridView1[colmn, row].Value != null)
                        {
                            int up = UpCut(), down = DownCut();
                            cutColumn = colmn;
                            cutRow = row;
                            cutCount = up + down+1;
                            cut_Buufer = arrayTable[row, colmn];
                            for (int i = row - up; i <= row + down; i++)
                            {
                                arrayTable[i, colmn]._disiplines = null;
                                arrayTable[i, colmn]._rooms = null;
                                arrayTable[i, colmn]._teacher = null;
                                arrayTable[i, colmn]._tema = null;
                                dataGridView1[colmn, i].Value = null;
                                dataGridView1[colmn, i].Style.BackColor = Color.White;
                                allCheakPozizitions(i, colmn);
                            }
                            button9.BackgroundImage = Properties.Resources.image2;
                            _pasteOrCut = true;

                        }
                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, запоните для начала ячейку");
                    }
                }
                else
                {
                    if (cutCount != 1)
                    {
                        if (daysAccept(row, cutColumn, cutCount-1))
                        {
                            if (DragAndDrop_chaekPool(row, colmn, 0, cutCount-1))
                            {
                                for (int i = row; i < row + cutCount; i++)
                                {
                                    arrayTable[i, colmn]._disiplines = cut_Buufer._disiplines;
                                    arrayTable[i, colmn]._teacher = cut_Buufer._teacher;
                                    arrayTable[i, colmn]._tema = cut_Buufer._tema;
                                    arrayTable[i, colmn]._rooms = cut_Buufer._rooms;
                                    
                                    dataGridView1[colmn, i].Value = cut_Buufer._disiplines + " " + cut_Buufer._tema + " " + cut_Buufer._rooms;
                                    foreach (var item in cut_Buufer._teacher)
                                    {
                                        dataGridView1[colmn, i].Value += " ";
                                        dataGridView1[colmn, i].Value += item;
                                        
                                    }
                                    allCheakPozizitions(i, colmn);
                                }
                                _pasteOrCut = false;
                                button9.BackgroundImage = Properties.Resources.cut_105155;
                            }
                            else
                            {
                                MessageBox.Show("Нехватает места для вставки данной пары");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Вы пытаетесь разорвать одну тему на несколько дней");
                        }

                    }
                    else
                    {
                        CellsTable b = arrayTable[row, colmn];
                        arrayTable[row, colmn]._disiplines = cut_Buufer._disiplines;
                        arrayTable[row, colmn]._tema = cut_Buufer._tema;
                        arrayTable[row, colmn]._teacher = cut_Buufer._teacher;
                        arrayTable[row, colmn]._rooms = cut_Buufer._rooms;
                        arrayTable[cutRow, cutColumn]._disiplines = b._disiplines;
                        arrayTable[cutRow, cutColumn]._tema = b._tema;
                        arrayTable[cutRow, cutColumn]._rooms = b._rooms;
                        arrayTable[cutRow, cutColumn]._teacher = b. _teacher;
                        dataGridView1[colmn, row].Value = cut_Buufer._disiplines + " " + cut_Buufer._tema + " " + cut_Buufer._rooms;
                        foreach (var item in cut_Buufer._teacher)
                        {
                            dataGridView1[colmn, row].Value += " ";
                            dataGridView1[colmn, row].Value += item;

                        }
                        dataGridView1[cutColumn, cutRow].Value = b._disiplines + " " + b._tema + " " + b._rooms + " " + b._teacher;

                        allCheakPozizitions(row, colmn);
                        allCheakPozizitions(cutRow, cutColumn);

                        _pasteOrCut = false;
                        button9.BackgroundImage = Properties.Resources.cut_105155;
                    }
                }
            }
            ColorDataGrid();
            VOSKRESENIE();
            ColorDataGridAll();
        }

        bool daysAccept(int frow, int fcolumn, int down)
        {
            if (down == 0)
            {
                return true;
            }

            int fdown = 0, i = 0;
            while (frow + i <= allRow && (dataGridView1[0, frow + i].Value.ToString() != "" || dataGridView1[0, frow + i + 1].Value.ToString() != ""))
            {
                fdown++;
                i++;
            }

            if (fdown < down)
                return false;

            return true;
        }
        
        // Drag and drop
        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            int SourceRow = Convert.ToInt32(e.Data.GetData(Type.GetType("System.Int32")));
            System.Drawing.Point clientPoint = dataGridView1.PointToClient(new System.Drawing.Point(e.X, e.Y));
            DataGridView.HitTestInfo hit = dataGridView1.HitTest(clientPoint.X, clientPoint.Y);
            int up = DragAndDrop_verificationsUp();
            int down = DragAndDrop_verificationsDown();
            numberGroup = dataGridView1[DragColumn, 3].Value.ToString();
            disciplines = arrayTable[DragRow, DragColumn]._disiplines;
            tema = arrayTable[DragRow, DragColumn]._tema;
            room = arrayTable[DragRow, DragColumn]._rooms;
            teach = arrayTable[DragRow, DragColumn]._teacher;

            if (dataGridView1[hit.ColumnIndex, hit.RowIndex].ReadOnly != true)
            {
                if (arrayTable[DragRow, DragColumn]._teacher != null)
                {
                    if (dataGridView1[hit.ColumnIndex, 3].Value.ToString() == dataGridView1[DragColumn, 3].Value.ToString())
                    {
                        if (daysAccept(hit.RowIndex, hit.ColumnIndex, up + down))
                        {
                            if ((up == 0 && down == 0) || DragAndDrop_chaekPool(hit.RowIndex, hit.ColumnIndex, 0, down + up))
                            {
                                if (dataGridView1[hit.ColumnIndex, hit.RowIndex].Value != null)
                                {

                                    DragAndDrop_change1(hit.ColumnIndex, hit.RowIndex, up, down);

                                    for (int i = 0; i <= down + up; i++)
                                    {
                                        allCheakPozizitions(DragRow - up + i, DragColumn);
                                        allCheakPozizitions(hit.RowIndex - up + i, hit.ColumnIndex);
                                    }
                                }
                                else
                                {
                                    DragAndDrop_change2(hit.ColumnIndex, hit.RowIndex, up, down);


                                    for (int i = 0; i <= down + up; i++)
                                    {
                                        allCheakPozizitions(DragRow - up + i, DragColumn);
                                        allCheakPozizitions(hit.RowIndex - up + i, hit.ColumnIndex);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Нехватает места для вставки данной пары");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Вы пытаетесь разорвать одну тему на несколько дней");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Вы пытаетесь вставить пары из группы " + dataGridView1[hit.ColumnIndex, 3].Value + " в группу " + dataGridView1[DragColumn, 3].Value);
                    }
                }
                else
                {
                    MessageBox.Show("Пожалуйста, запоните для начала ячейку");
                }
            }
            ColorDataGrid();
            VOSKRESENIE();
            ColorDataGridAll();
        }

        public bool DragAndDrop_chaekPool(int Frow, int Fcol, int up, int down)
        {
            for (int i = Frow - up; i <= Frow + down; i++)
            {
                if (arrayTable[i, Fcol]._disiplines != null && arrayTable[i, Fcol]._disiplines != "")
                    return false;
            }

            return true;
        }

        public void DragAndDrop_change1(int ColumnIndex, int RowIndex, int up, int down)
        {
            int RowTwo = DragRow - up;
            for (int i = RowIndex; i <= RowIndex + down + up; i++)
            {
                string[] teachBuffer;
                string bufersDataGrid = dataGridView1[ColumnIndex, i].Value.ToString();
                dataGridView1[ColumnIndex, i].Value = dataGridView1[DragColumn, RowTwo].Value;
                dataGridView1[DragColumn, RowTwo].Value = bufersDataGrid;

                bufersDataGrid = arrayTable[i, ColumnIndex]._disiplines;
                arrayTable[i, ColumnIndex]._disiplines = arrayTable[RowTwo, DragColumn]._disiplines;
                arrayTable[RowTwo, DragColumn]._disiplines = bufersDataGrid;

                bufersDataGrid = arrayTable[i, ColumnIndex]._rooms;
                arrayTable[i, ColumnIndex]._rooms = arrayTable[RowTwo, DragColumn]._rooms;
                arrayTable[RowTwo, DragColumn]._rooms = bufersDataGrid;

                teachBuffer = arrayTable[i, ColumnIndex]._teacher;
                arrayTable[i, ColumnIndex]._teacher = arrayTable[RowTwo, DragColumn]._teacher;
                arrayTable[RowTwo, DragColumn]._teacher = teachBuffer;

                bufersDataGrid = arrayTable[i, ColumnIndex]._tema;
                arrayTable[i, ColumnIndex]._tema = arrayTable[RowTwo, DragColumn]._tema;
                arrayTable[RowTwo, DragColumn]._tema = bufersDataGrid;
                RowTwo++;
            }

        }

        public void DragAndDrop_change2(int ColumnIndex, int RowIndex, int up, int down)
        {
            int RowTwo = DragRow - up;
            for (int i = RowIndex; i <= RowIndex + down + up; i++)
            {
                dataGridView1[ColumnIndex, i].Value = dataGridView1[DragColumn, RowTwo].Value;
                dataGridView1[DragColumn, RowTwo].Value = null;


                arrayTable[i, ColumnIndex]._disiplines = arrayTable[RowTwo, DragColumn]._disiplines;
                arrayTable[RowTwo, DragColumn]._disiplines = "";

                arrayTable[i, ColumnIndex]._rooms = arrayTable[RowTwo, DragColumn]._rooms;
                arrayTable[RowTwo, DragColumn]._rooms = "";

                arrayTable[i, ColumnIndex]._teacher = arrayTable[RowTwo, DragColumn]._teacher;
                arrayTable[RowTwo, DragColumn]._teacher = null;


                arrayTable[i, ColumnIndex]._tema = arrayTable[RowTwo, DragColumn]._tema;
                arrayTable[RowTwo, DragColumn]._tema = "";

                dataGridView1[DragColumn, RowTwo].Style.BackColor = Color.White;
                RowTwo++;
            }
        }

        public int DragAndDrop_verificationsUp()
        {
            int answer = 0;

            for (int i = DragRow - 3; i < DragRow; i++)
                if (arrayTable[DragRow, DragColumn]._disiplines == arrayTable[i, DragColumn]._disiplines && arrayTable[DragRow, DragColumn]._tema == arrayTable[i, DragColumn]._tema)
                    answer++;

            return answer;
        }

        public int UpCut()
        {
            int answer = 0;

            for (int i = row - 3; i < row; i++)
                if (arrayTable[row, colmn]._disiplines == arrayTable[i, colmn]._disiplines && arrayTable[row, colmn]._tema == arrayTable[i, colmn]._tema)
                    answer++;

            return answer;
        }
        public int DragAndDrop_verificationsDown()
        {
            int answer = 0;

            for (int i = DragRow + 3; i > DragRow; i--)
                if (arrayTable[DragRow, DragColumn]._disiplines == arrayTable[i, DragColumn]._disiplines && arrayTable[DragRow, DragColumn]._tema == arrayTable[i, DragColumn]._tema)
                    answer++;

            return answer;
        }
        public int DownCut()
        {
            int answer = 0;

            for (int i = row + 3; i > row; i--)
                if (arrayTable[row, colmn]._disiplines == arrayTable[i, colmn]._disiplines && arrayTable[row, colmn]._tema == arrayTable[i, colmn]._tema)
                    answer++;

            return answer;
        }

        // Поиск группы
        public int searchGroup(string name)
        {
            int countForARray = 0;
            for (int i = 0; i < lessons.Length; i++)
            {
                if (lessons[i] != null && lessons[i]._nameAllLessin == name)
                {
                    break;
                }
                countForARray++;
            }

            return countForARray;
        }

        // Поиск номера группы
        public string searchStringGroup(int fColmn)
        {
            return dataGridView1[fColmn, 3].Value.ToString();
        }

        // Получение номера дисциплины в учительском массиве
        //public int GetNumberDisciplinesForTeacherArray(Teacher teacher, int numbergroup, string disciplens)
        //{
        //    int count = 0;
        //    for (int i = 0; i < teacher.getCount(); i++)
        //    {
        //        if (teacher._arrayGroup[numbergroup]._arrayDiscipline[i]._nameDiscipline == disciplens)
        //        {
        //            count = i;
        //            return count;
        //        }
        //    }
        //    return count;
        //}

        //Вывод ошибок
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Error happened " + e.Context.ToString());
        }

        //Удалить
        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("3");
        }

        //Событие нажатия на ячейку 
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label2.Text = "";
            row = e.RowIndex;
            colmn = e.ColumnIndex;
            int countCheakDisiplines = 0;
            if (row > -1 && colmn > -1)
            {
                
                numberGroup = dataGridView1[colmn, 3].Value.ToString();

                if (checkBox1.Checked)
                {
                    
                    if (!(arrayTable[row, colmn]._disiplines == null && arrayTable[row, colmn]._disiplines == ""))
                    {
                        if (cheakCells(arrayTable[e.RowIndex, e.ColumnIndex]))
                        {
                            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null && dataGridView1[e.ColumnIndex, e.RowIndex].ReadOnly != true && checkBox1.Checked)
                            {
                                label2.Text = searchCommentsTems(numberGroup, arrayTable[e.RowIndex, e.ColumnIndex]._disiplines, arrayTable[e.RowIndex, e.ColumnIndex]._tema);
                                Form4 form = new Form4(arrayTable[e.RowIndex, e.ColumnIndex]._disiplines, arrayTable[e.RowIndex, e.ColumnIndex]._tema, ref arrayTable, ref sender, ref e, clas, teacher, dataGridView1[e.ColumnIndex, 3].Value.ToString(), ref dataGridView1, dataGridView1[e.ColumnIndex,4].Value.ToString());
                                form.ShowDialog();
                                countCheakDisiplines = form.GetDistance;
                            }
                        }
                        else
                        {
                            if (dataGridView1[e.ColumnIndex, e.RowIndex].Value != null && dataGridView1[e.ColumnIndex, e.RowIndex].ReadOnly != true && checkBox1.Checked)
                            {
                                label2.Text = searchCommentsTems(numberGroup, arrayTable[e.RowIndex, e.ColumnIndex]._disiplines, arrayTable[e.RowIndex, e.ColumnIndex]._tema);
                                Form4 form = new Form4( arrayTable[e.RowIndex, e.ColumnIndex]._disiplines, arrayTable[e.RowIndex, e.ColumnIndex]._tema, arrayTable[e.RowIndex, e.ColumnIndex]._teacher, arrayTable[e.RowIndex, e.ColumnIndex]._rooms, ref arrayTable, ref sender, ref e, clas, teacher, dataGridView1[e.ColumnIndex, 3].Value.ToString(), ref dataGridView1, dataGridView1[e.ColumnIndex, 4].Value.ToString());
                                form.ShowDialog();
                                countCheakDisiplines = form.GetDistance;
                            }
                        }
                        for (int i = row-countCheakDisiplines ; i <= row - countCheakDisiplines + countCheakDisiplines * 2; i++) 
                        {
                            allCheakPozizitions(i, colmn);
                        }
                    }
                }
            }
            
        }

        public string searchCommentsTems(string nameGroupe,string nameDisiplines, string tems)
        {
            int countSearchForGroup = 0, countSearchTems = 0;
            while (countSearchForGroup < lessons.Length && lessons[countSearchForGroup]._nameAllLessin != nameGroupe)
                countSearchForGroup++;

            while (countSearchTems < lessons[countSearchForGroup]._arrayLesson.Length && lessons[countSearchForGroup]._arrayLesson[countSearchTems]._tema != tems)
                countSearchTems++;
            return lessons[countSearchForGroup]._arrayLesson[countSearchTems]._coments;
        }

        public bool cheakCells(CellsTable cell)
        {
            if (cell._disiplines != null && cell._disiplines != "" && cell._tema != null && cell._tema != "" && (cell._rooms == null || cell._rooms == "") && (cell._teacher == null || cell._teacher == null
                ))
                return true;

            return false;
        }

        //Поиск темы
        public int searchTems(string fDisciplines, string fTems)
        {
            int countless = searchGroup(numberGroup);
            for (int i = 0; i < lessons[countless]._arrayLesson.Length; i++)
            {
                if (lessons[countless]._arrayLesson[i]._nameDiscipline == fDisciplines && lessons[countless]._arrayLesson[i]._tema == fTems)
                {
                    return i;
                }
            }
            return -1;
        }

        //Закрашивает воскресенье
        public void VOSKRESENIE()
        {
            for (int i = 7; i < allRow; i += 4)
            {
                if (dataGridView1[0, i].Value.ToString() == "вс")
                {
                    for (int j = 0; j < allColumn; j++)
                    {
                        dataGridView1[j, i - 2].Style.BackColor = Color.Orange;
                        dataGridView1[j, i - 1].Style.BackColor = Color.Orange;
                        dataGridView1[j, i].Style.BackColor = Color.Orange;
                        dataGridView1[j, i + 1].Style.BackColor = Color.Orange;
                        arrayTable[i - 2, j]._down = true;
                        arrayTable[i - 1, j]._down = true;
                        arrayTable[i, j]._down = true;
                        arrayTable[i + 1, j]._down = true;
                    }
                }
            }
        }

        //Проверка на универсальность дисциплины
        public bool CheakInuversalDisiplines(AllLessin[] less, int count1, int count2)
        {
            for (int i = 0; i < count1; i++)
            {
                for (int j = 0; j < less[i]._arrayLesson.Length; j++)
                {
                    if (less[i]._arrayLesson[j]._nameDiscipline == less[count1]._arrayLesson[count2]._nameDiscipline)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        //Получение числа универсальных дисциплин
        public int GetCountUniversalDisiplines(AllLessin[] less)
        {
            int count = 0;
            for (int i = 0; i < less.Length; i++)
            {
                if (less[i] != null)
                    for (int j = 0; j < less[i]._arrayLesson.Length; j++)
                    {
                        if (CheakInuversalDisiplines(less, i, j))
                            count++;
                    }

            }


            return count;
        }

        //Цвет 
        public void ColorDataGrid()
        {
            for (int i = 0; i < allRow; i++)
            {
                for (int j = 0; j < allColumn; j++)
                {
                    if (arrayTable[i, j]._down == true)
                    {
                        dataGridView1[j, i].Style.BackColor = Color.Gray;
                    }
                }
            }
        }

        public void ColorDataGridAll()
        {
            for (int i = 0; i < allRow; i++)
                for (int j = 0; j < allColumn; j++)
                {
                    if (arrayTable[i, j]._disiplines != null)
                    {
                        int numberArr = arrayColorDisiplines.SearhDisiplines(arrayTable[i, j]._disiplines);
                        if (numberArr != -1)
                        {
                            dataGridView1[j, i].Style.BackColor = arrayColorDisiplines._arrayColorDisiplines[numberArr]._color;
                        }
                    }
                }
        }

        //Глобалка
        public void getAlllogical()
        {
            string buffClass = "";
            string [] buufTeachFB = new string[0];
            int numberGroup;
            for (int count = 0; count < allColumn; count++)
            {
                if (arrayTable[1, count]._down == true)
                {
                    int countRow = 5;
                    numberGroup = searchGroup(dataGridView1[count, 3].Value.ToString());
                    for (int i = 0; i < lessons[numberGroup]._arrayLesson.Length; i++)
                    {
                        int countHour = Convert.ToInt32(lessons[numberGroup]._arrayLesson[i]._time);

                        while (countHour > 0)
                        {
                            dataGridView1[count, countRow].Value = lessons[numberGroup]._arrayLesson[i]._nameDiscipline + " " + lessons[numberGroup]._arrayLesson[i]._tema;
                            arrayTable[countRow, count]._disiplines = lessons[numberGroup]._arrayLesson[i]._nameDiscipline;
                            arrayTable[countRow, count]._tema = lessons[numberGroup]._arrayLesson[i]._tema;
                            if (countRoomsFB(lessons[numberGroup]._arrayLesson[i]._nameDiscipline, dataGridView1[count, 3].Value.ToString(), lessons[numberGroup]._arrayLesson[i]._tema,ref buffClass) == 1)
                            {
                                arrayTable[countRow, count]._rooms = buffClass;
                                dataGridView1[count, countRow].Value = dataGridView1[count, countRow].Value + " " + buffClass;
                            }
                            
                            
                            if (countPrepodsFB (lessons[numberGroup]._arrayLesson[i]._nameDiscipline, dataGridView1[count, 3].Value.ToString(), lessons[numberGroup]._arrayLesson[i]._tema, ref buufTeachFB) ==1)
                            {
                                arrayTable[countRow, count]._teacher = buufTeachFB;
                                dataGridView1[count, countRow].Value = dataGridView1[count, countRow].Value + " " + buufTeachFB[0];
                            }
                            countRow++;
                            countHour -= 2;
                        }
                    }
                }
            }
        }
        //Проверка на преподов
        public int countPrepodsFB (string nameDisiplines, string numberGroupe, string teme, ref string[] buffTeach)
        {
            buffTeach = new string[1];
            for (int i = 0; i < teacher._array.Length;i++)
            {
                if (teacher._array[i]._nameDisiplines == nameDisiplines && teacher._array[i]._numberGroupe == numberGroupe)
                {
                    if (teacher._array[i]._teachers.Length == 1)
                        buffTeach = teacher._array[i]._teachers;
                    return teacher._array[i]._teachers.Length;
                }
            }

            return -1;
        }
        public int countRoomsFB(string nameDisiplines, string numberGroupe, string teme, ref string buffClass)
        {
            bool myCheakDead = false;
            for (int i = 0; i < teacher._array.Length; i++)
            {
                if (clas._array[i].numberGroupe == numberGroupe && clas._array[i].AllLessin == nameDisiplines)
                {
                    if (!myCheakDead)
                    {
                        buffClass = clas._array[i].rooms;
                        myCheakDead = true;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            return 1;
        }



        //Проверка по горизонтали
        public void cheakGorizontal(int fRow)
        {
            for (int j = 0; j < allColumn; j++)
            {
                if (arrayTable[fRow, j]._disiplines != null)
                {
                    for (int i = 0; i < allColumn; i++)
                    {
                        if (dataGridView1[i, fRow].Value != null && arrayTable[fRow, i]._down != true && arrayTable[fRow, j]._down != true)
                        {
                            if (arrayTable[fRow, j]._rooms == arrayTable[fRow, i]._rooms && j != i)
                            {
                                dataGridView1[j, fRow].Style.ForeColor = colorDialog2.Color;
                                dataGridView1[i, fRow].Style.ForeColor = colorDialog2.Color;
                                if (dataGridView1[j, fRow].ToolTipText.IndexOf("У вас совпали классы в данный день группе " + searchStringGroup(i) + " и " + searchStringGroup(j) + "\n") == -1)
                                    dataGridView1[j, fRow].ToolTipText += "У вас совпали классы в данный день группе " + searchStringGroup(i) + " и " + searchStringGroup(j) + "\n";
                            }
                            
                            if (arrayTable[fRow, j]._teacher !=  null && arrayTable[fRow, i]._teacher != null && EqualArrTeachers(arrayTable[fRow, j]._teacher, arrayTable[fRow, i]._teacher)>0 && j != i && arrayTable[fRow, i]._down != true && arrayTable[fRow, j]._down != true)
                            {
                                dataGridView1[j, fRow].Style.ForeColor = colorDialog1.Color;
                                dataGridView1[i, fRow].Style.ForeColor = colorDialog1.Color;
                                if (dataGridView1[j, fRow].ToolTipText.IndexOf("У вас совпали преподаватели в данный день группе " + searchStringGroup(i) + " и " + searchStringGroup(j) + "\n") == -1)
                                    dataGridView1[j, fRow].ToolTipText += "У вас совпали преподаватели в данный день группе " + searchStringGroup(i) + " и " + searchStringGroup(j) + "\n";
                            }
                        }
                    }
                }
            }
        }

        //Проверка двух массивов
        int EqualArrTeachers (string [] arr1, string [] arr2)
        {
            int count = 0;

            foreach (var itemsArray1 in arr1)
            {
                foreach (var itemsArray2 in arr2)
                {
                    if (itemsArray1 == itemsArray2)
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        //Првоерка по вертикали
        public void cheakVertical(int Fcolumn)
        {
            for (int j = 0; j < allRow; j++)
            {
                if (arrayTable[j, Fcolumn]._disiplines != null)
                {
                    for (int i = 0; i < j; i++)
                    { 
                        if (arrayTable[i, Fcolumn]._disiplines != null && arrayTable[i, Fcolumn]._disiplines.Length != 0 && arrayTable[j, Fcolumn]._disiplines != null && arrayTable[j, Fcolumn]._disiplines.Length != 0 && searchTems(arrayTable[i, Fcolumn]._disiplines, arrayTable[i, Fcolumn]._tema) > searchTems(arrayTable[j, Fcolumn]._disiplines, arrayTable[j, Fcolumn]._tema) && arrayTable[i, Fcolumn]._down != true && arrayTable[j, Fcolumn]._down != true && arrayTable[i, Fcolumn]._disiplines == arrayTable[j, Fcolumn]._disiplines)
                        {
                            dataGridView1[Fcolumn, j].Style.ForeColor = colorDialog4.Color;
                            dataGridView1[Fcolumn, i].Style.ForeColor = colorDialog4.Color;
                            if (dataGridView1[Fcolumn, j].ToolTipText.IndexOf("Нарушение построения логики" + "\n") == -1)
                                dataGridView1[Fcolumn, j].ToolTipText += "Нарушение построения логики" + "\n";
                            if (dataGridView1[Fcolumn, i].ToolTipText.IndexOf("Нарушение построения логики" + "\n") == -1)
                                dataGridView1[Fcolumn, i].ToolTipText += "Нарушение построения логики" + "\n";
                        }
                    }
                }
            }
        }

        //Проверка при вставке или замене элемента
        public void allCheakPozizitions(int frow, int fcolumn)
        {
            for (int i = 0; i < allRow; i++)
            {
                dataGridView1[fcolumn, i].Style.ForeColor = Color.Black;
                dataGridView1[fcolumn, i].ToolTipText = "";
            }

            for (int i = 0; i < allColumn; i++)
            {
                dataGridView1[i, frow].Style.ForeColor = Color.Black;
                dataGridView1[i, frow].ToolTipText = "";
            }

            for (int i = 0; i < allRow; i++)
            {
                cheakGorizontal(i);
            }

            cheakVertical(fcolumn);

        }
    }

    public struct Tema
    {
        public string _tema;
        public string[] _teachers;
        public string[] _rooms;
    }
    public struct Discipline
    {
        public string _nameDiscipline;
        public Tema[] _tems;
        public int _countTems;
    }
    public class OneAllLessin
    {
        public string _nameAllLessin;
        public int _countDiscipline;
        public Discipline[] _disciplines;

        public OneAllLessin(int countdiscipline)
        {
            _disciplines = new Discipline[countdiscipline];
            _countDiscipline = 0;
        }

        ~OneAllLessin()
        {

        }

        /// <summary>
        /// Добавляет дисциплину
        /// </summary>
        /// <param name="countTheme">Количество тем в дисциплине</param>
        /// <param name="discipline">Дисциплина</param>
        public void pushDiscipline(Discipline discipline)
        {
            _disciplines[_countDiscipline] = discipline;
            _countDiscipline++;
        }

    }
    public class OneGroup
    {
        public Discipline[] _disciplines;
        public string _numberGroup;

        public OneGroup(string numberGroup, int countDisciplines)
        {
            _disciplines = new Discipline[countDisciplines];
            _numberGroup = numberGroup;
        }
        public void PushDiscipline(Discipline[] e)
        {
            _disciplines = e;
        }
    }
    public class AllGroup
    {
        public OneGroup[] _group;
        private int _count;

        public AllGroup(int countGroup)
        {
            _count = 0;
            _group = new OneGroup[countGroup];
        }
        public void PushOneGroup(OneGroup e)
        {
            _group[_count] = e;
            _count++;
        }
    }
    public struct CellsTable
    {
        public string [] _teacher;
        public string _rooms;
        public string _disiplines;
        public string _tema;
        public bool _down;
    }
    public struct ColorDisiplines
    {
        public Color _color;
        public string _disiplines;
    }
    public class AllColorDisipilens
    {
        public ColorDisiplines[] _arrayColorDisiplines;
        public int _lenght;
        public int _count;
        public AllColorDisipilens(int count)
        {
            _count = 0;
            _arrayColorDisiplines = new ColorDisiplines[count];
            _lenght = count;
        }
        public void PushColorDisiplines(Color color, string disi)
        {
            int numberColor = SearhDisiplines(disi);
            if (SearhDisiplines(disi) == -1)
            {
                _arrayColorDisiplines[_count]._color = color;
                _arrayColorDisiplines[_count]._disiplines = disi;
                _count++;
            }
            else
            {
                _arrayColorDisiplines[numberColor]._color = color;
            }

        }
        public int SearhDisiplines(string disi)
        {
            for (int i = 0; i < _count; i++)
            {
                if (_arrayColorDisiplines[i]._disiplines == disi)
                    return i;
            }

            return -1;
        }


    }
    public struct OnePozitionsTeacher
    {
        public string _nameDis;
        public string _nameGroup;
        public string _teacher;
        public string _team;
    }
    public class TeacherForSql
    {
        public OnePozitionsTeacher[] _array;
        public int _count;

        public TeacherForSql()
        {
            _array = new OnePozitionsTeacher[0];
            _count = 0;
        }

        public void push(string nameDis, string nameGroupe, string team, string teacher)
        {
            OnePozitionsTeacher[] answer = new OnePozitionsTeacher[_count + 1];
            for (int i = 0; i < _array.Length; i++)
            {
                answer[i] = _array[i];
            }
            answer[_count]._nameDis = nameDis;
            answer[_count]._nameGroup = nameGroupe;
            answer[_count]._teacher = teacher;
            answer[_count]._team = team;
            _count++;
            _array = answer;
        }
    }

}
