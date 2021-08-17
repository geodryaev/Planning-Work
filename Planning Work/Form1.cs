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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DateTime date = new DateTime();
            if (!(Properties.Settings.Default.DateForMe == date))
            {
                dateTimePicker1.Value = Properties.Settings.Default.DateForMe;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        int countUploadFile = 0; //Количество заруженых сетевых графикоф
        Allpeople baseItem = new Allpeople();
        AllLessin[] allLesson = new AllLessin[5000];
        Teacher teacher;
        int countAllLessons = 0;
        int Kastil1 = 0;
        
        KMA clas = new KMA(); 
        AllLessinAndRooms AllLessinRooms = new AllLessinAndRooms();

        //time
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DateForMe = dateTimePicker1.Value;
            Properties.Settings.Default.Save();
        }

        //-----------------------------------Форма для расписания
        private void button7_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2( countUploadFile, dateTimePicker1.Value);
            form.ShowDialog();
            
        }
        //-----------------------------------Добавить таблицу в базу данных
        private void button1_Click(object sender, EventArgs e)
        {
            string pathFile = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathFile = openFileDialog1.FileName;
            

                //открытие книги
                if (ItsExsel(pathFile))
                {

                    Excel.Application App;
                    Excel.Workbook xlsWB;
                    Excel.Worksheet xlsSheet;
                    

                    App = new Excel.Application();
                    xlsWB = App.Workbooks.Open(@pathFile);
                    xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);

                    int countWorksheets = xlsWB.Worksheets.Count;

                    progressBar1.Maximum = countWorksheets*2;
                    progressBar1.Value = 0;
                    for (int numberPage = 1; numberPage <= countWorksheets; numberPage++)
                    {
                        
                        xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(numberPage);

                        int count = 0 ;
                        string name = xlsSheet.Name;

                        int[] arrayAllLessin = getAllLessin(xlsSheet.Cells[1, 1].Text, ref count);
                        int countTriple = 0;
                        Triple[] lesson = getLess(xlsSheet, ref countTriple);
                        Kastil1++;
                        progressBar1.Value++;
                        for (int i = 0; i < count; i++) 
                        {
                            baseItem.pushElements(name, arrayAllLessin[i], NormalNameKyrs(xlsWB.Name));
                            allLesson[countAllLessons] = new AllLessin(Convert.ToString(arrayAllLessin[i]),countTriple);
                            allLesson[countAllLessons].pushtriple(lesson);
                            countAllLessons++;
                        }
                        progressBar1.Value++;
                    }

                    countUploadFile++;
                    xlsWB.Save();
                    xlsWB.Close(true);
                    App.Quit();
                    progressBar1.Value = 0;
                }
                else
                {
                    MessageBox.Show("Выбраный файл не является Excel");
                }
            }
        }

        //-----------------------------------Создание таблицы
        private void button2_Click(object sender, EventArgs e)
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


                int ch = 0;
                int num = 3;
               
                for (int i = 0; i< baseItem._count;)
                {
                    createOneTemplate(ref xlsSheet, ch, num);
                    setAllLessin(ref xlsSheet, ref ch, ref num, baseItem, ref i);
                }
                xlsWB.SaveAs(pathFile);
                xlsWB.Close(true);
                App.Quit();
            }
        }

        //-----------------------------------Добавление преподавателей
        private void button4_Click(object sender, EventArgs e)
        {

            string pathFile = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathFile = openFileDialog1.FileName;


                //открытие книги
                if (ItsExsel(pathFile))
                {

                    Excel.Application App;
                    Excel.Workbook xlsWB;
                    Excel.Worksheet xlsSheet;


                    App = new Excel.Application();
                    xlsWB = App.Workbooks.Open(@pathFile);
                    xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);

                    
                    int countWorksheets = xlsWB.Worksheets.Count;
                    int count=0;
                    for (int numberPage = 1; numberPage <= countWorksheets; numberPage++)
                    {
                        xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(numberPage);
                        count += countGroupForTeacher(xlsSheet);
                    }
                    teacher = new Teacher();
                    for (int numberPage = 1; numberPage <= countWorksheets; numberPage++)
                    {
                        xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(numberPage);
                        getGoodMan(xlsSheet);
                    }
                    
                    xlsWB.Save();
                    xlsWB.Close(true);
                    App.Quit();
                }
                else
                {
                    MessageBox.Show("Выбраный файл не является Excel");
                }
            }
        }
        
        //-----------------------------------Очистка памяти
        private void button3_Click(object sender, EventArgs e)
        {

        }

        //-----------------------------------Крепление классов и предметов
        private void button5_Click(object sender, EventArgs e)
        {
            string pathFile = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathFile = openFileDialog1.FileName;


                //открытие книги
                if (ItsExsel(pathFile))
                {

                    Excel.Application App;
                    Excel.Workbook xlsWB;
                    Excel.Worksheet xlsSheet;


                    App = new Excel.Application();
                    xlsWB = App.Workbooks.Open(@pathFile);
                    xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);

                    PKP(xlsSheet);
                    xlsWB.Save();
                    xlsWB.Close(true);
                    App.Quit();
                }
                else
                {
                    MessageBox.Show("Выбраный файл не является Excel");
                }
            }
        }

        //-----------------------------------Приязка групп и классов
        private void button6_Click(object sender, EventArgs e)
        {
            string pathFile = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathFile = openFileDialog1.FileName;


                //открытие книги
                if (ItsExsel(pathFile))
                {

                    Excel.Application App;
                    Excel.Workbook xlsWB;
                    Excel.Worksheet xlsSheet;


                    App = new Excel.Application();
                    xlsWB = App.Workbooks.Open(@pathFile);
                    xlsSheet = (Excel.Worksheet)xlsWB.Worksheets.get_Item(1);


                    getGP(xlsSheet);
                    xlsWB.Save();
                    xlsWB.Close(true);
                    App.Quit();
                }
                else
                {
                    MessageBox.Show("Выбраный файл не является Excel");
                }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            SqlDataReader buff;
            bool Dont_createTable;
            using (var connection = new SqlConnection(get_cs()))
            {
                //Для Allpeople
                {
                    connection.Open();
                    using (var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME='_allPeolpe';", connection))
                    {
                        buff = cmd.ExecuteReader();
                        Dont_createTable = buff.HasRows;
                    }
                    connection.Close();
                    connection.Open();
                    if (Dont_createTable == false)
                    {
                        using (var cmd = new SqlCommand("CREATE TABLE _allPeolpe ( ID int NOT NULL IDENTITY(1,1) primary key, nameGroup NVARCHAR(MAX), fac NVARCHAR(MAX), nameDisciplines NVARCHAR(MAX), commetsDisciplines NVARCHAR(MAX), tema NVARCHAR(MAX), comments NVARCHAR(MAX), timeLection NVARCHAR(MAX), setA NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SqlCommand("DROP TABLE _allPeolpe; \n CREATE TABLE _allPeolpe ( ID int NOT NULL IDENTITY(1,1) primary key, nameGroup NVARCHAR(MAX), fac NVARCHAR(MAX), nameDisciplines NVARCHAR(MAX), commetsDisciplines NVARCHAR(MAX), tema NVARCHAR(MAX), comments NVARCHAR(MAX), timeLection NVARCHAR(MAX), setA NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }
                //Для teacher
                {
                    connection.Open();
                    using (var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME='_teacher';", connection))
                    {
                        buff = cmd.ExecuteReader();
                        Dont_createTable = buff.HasRows;
                    }
                    connection.Close();
                    connection.Open();
                    if (Dont_createTable == false)
                    {
                        using (var cmd = new SqlCommand("CREATE TABLE _teacher ( ID int NOT NULL IDENTITY(1,1) primary key, nameDisciplines NVARCHAR(MAX),  nameGroup NVARCHAR(MAX),team NVARCHAR(MAX) ,teacher NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SqlCommand("DROP TABLE _teacher; \n CREATE TABLE _teacher ( ID int NOT NULL IDENTITY(1,1) primary key, nameDisciplines NVARCHAR(MAX),  nameGroup NVARCHAR(MAX),team NVARCHAR(MAX) ,teacher NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }
                //Для class
                {
                    connection.Open();
                    using (var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME='_class';", connection))
                    {
                        buff = cmd.ExecuteReader();
                        Dont_createTable = buff.HasRows;
                    }
                    connection.Close();
                    connection.Open();
                    if (Dont_createTable == false)
                    {
                        using (var cmd = new SqlCommand("CREATE TABLE _class ( ID int NOT NULL IDENTITY(1,1) primary key, name NVARCHAR(MAX),  room NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SqlCommand("DROP TABLE _class; \n CREATE TABLE _class ( ID int NOT NULL IDENTITY(1,1) primary key, name NVARCHAR(MAX),  room NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }
                //Для baseItems
                {
                    connection.Open();
                    using (var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME='_baseItems';", connection))
                    {
                        buff = cmd.ExecuteReader();
                        Dont_createTable = buff.HasRows;
                    }
                    connection.Close();
                    connection.Open();
                    if (Dont_createTable == false)
                    {
                        using (var cmd = new SqlCommand("CREATE TABLE _baseItems ( ID int NOT NULL IDENTITY(1,1) primary key,  AllLessing NVARCHAR(MAX), class NVARCHAR(MAX), fac NVARCHAR(MAX), kyrs NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SqlCommand("DROP TABLE _baseItems; \n CREATE TABLE _baseItems ( ID int NOT NULL IDENTITY(1,1) primary key,  AllLessing NVARCHAR(MAX), class NVARCHAR(MAX), fac NVARCHAR(MAX), kyrs NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }
                //Для Other
                {
                    connection.Open();
                    using (var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE  TABLE_NAME='_other';", connection))
                    {
                        buff = cmd.ExecuteReader();
                        Dont_createTable = buff.HasRows;
                    }
                    connection.Close();
                    connection.Open();
                    if (Dont_createTable == false)
                    {
                        using (var cmd = new SqlCommand("CREATE TABLE _other ( ID int NOT NULL IDENTITY(1,1) primary key, _row NVARCHAR(MAX), _column NVARCHAR(MAX), _countGroupe NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (var cmd = new SqlCommand("DROP TABLE _other; \n CREATE TABLE _other ( ID int NOT NULL IDENTITY(1,1) primary key, _row NVARCHAR(MAX), _column NVARCHAR(MAX), _countGroupe NVARCHAR(MAX))", connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    connection.Close();
                }

                
                for (int i = 0; i < baseItem._count; i++)
                {
                    for (int j = 0; j < allLesson[i]._arrayLesson.Length; j++)
                    {
                        string sqlComandForSetData = "INSERT INTO _allPeolpe (nameGroup, fac, nameDisciplines, commetsDisciplines, tema, comments, timeLection, setA) VALUES";
                        sqlComandForSetData += " (\'" + allLesson[i]._nameAllLessin.Trim() + "\', \'" + allLesson[i]._arrayLesson[j]._fack + "\', \'" + allLesson[i]._arrayLesson[j]._nameDiscipline.Trim()+ "\', \'" + allLesson[i]._arrayLesson[j]._comentsDisciplines.Trim() + "\', \'" + allLesson[i]._arrayLesson[j]._tema.Trim() + "\', \'" + allLesson[i]._arrayLesson[j]._coments.Trim() + "\', \'" + allLesson[i]._arrayLesson[j]._time.Trim() + "\', \'" + allLesson[i]._arrayLesson[j]._set + "\')";
                        connection.Open();
                        using (var cmd = new SqlCommand(sqlComandForSetData, connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                        connection.Close();
                    }
                }

                for (int i = 0; i < teacher._array.Length; i++)
                {
                    for (int j = 0; j < teacher._array[i]._teachers.Length;j++)
                    {
                        string sqlComandForSetData = "INSERT INTO _teacher (nameDisciplines, nameGroup, team, teacher) VALUES";
                        sqlComandForSetData += " (\'" + teacher._array[i]._nameDisiplines.Trim() + "\', \'" + teacher._array[i]._numberGroupe.Trim() + "\', \'" + teacher._array[i]._tems.Trim() + "\', \'" + teacher._array[i]._teachers[j].Trim() + "\')";
                        connection.Open();
                        using (var cmd = new SqlCommand(sqlComandForSetData, connection))
                        {
                            cmd.ExecuteNonQuery();
                        }
                        connection.Close();
                    }
                }

                for (int i = 0; i < AllLessinRooms._array.Length; i++)
                {
                    string sqlComandForSetData = "INSERT INTO _class (name, room) VALUES";
                    sqlComandForSetData += " (\'" + AllLessinRooms._array[i].AllLessin.Trim() + "\', \'" + AllLessinRooms._array[i].rooms.Trim() + "\')";
                    connection.Open();
                    using (var cmd = new SqlCommand(sqlComandForSetData, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    connection.Close();
                }

                for (int i = 0; i < baseItem._count; i++)
                {
                    string sqlComandForSetData = "INSERT INTO _baseItems (AllLessing, class, fac, kyrs) VALUES";
                    sqlComandForSetData += " (\'" + baseItem._all[i].AllLessin + "\', \'" + baseItem._all[i].clas.Trim() + "\', \'" + baseItem._all[i].fac.Trim() + "\', \'" + baseItem._all[i].kyrs.Trim() + "\')";
                    connection.Open();
                    using (var cmd = new SqlCommand(sqlComandForSetData, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    connection.Close();
                }

                {
                    string sqlComandForSetData = "INSERT INTO _other (_row, _column, _countGroupe) VALUES";
                    sqlComandForSetData += " (\'" +  CountRowForNewForms()  + "\', \'" +CountColumnForNewForms(baseItem._count, countUploadFile) + "\', \'" + baseItem._count + "\')";
                    connection.Open();
                    using (var cmd = new SqlCommand(sqlComandForSetData, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    connection.Close();
                }
            }

        }

        private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormSettingsForSqlServer form = new FormSettingsForSqlServer();
            form.ShowDialog();
        }

        //Нормальное название курса
        string NormalNameKyrs(string kyrs)
        {
            string answer = kyrs;
            int i = 0;
            
            while (kyrs[i] !='.')
            {
                i++;
            }
            
            return answer.Substring(0, answer.Length  - (answer.Length - i)); 
        }

        //Проверка файла
        bool ItsExsel(string path)
        {
            if (path != "")
            {
                int i = 0;
                while (path[i] != '.')
                    i++;

                char[] ext = new char[3];
                int count = 0;
                while (count < 3)
                {
                    i++;
                    ext[count] = path[i];
                    count++;
                }

                string exp = new string(ext);
                if (exp == "xls")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        //Русификация месяцев
        string RusMont(string engMont)
        {
            if (engMont == "1")
                return "январь";

            if (engMont == "2")
                return "февраль";

            if (engMont == "3")
                return "март";

            if (engMont == "4")
                return "апрель";

            if (engMont == "5")
                return "май";

            if (engMont == "6")
                return "июнь";

            if (engMont == "7")
                return "июль";

            if (engMont == "8")
                return "август";

            if (engMont == "9")
                return "сентябрь";

            if (engMont == "10")
                return "октябырь";

            if (engMont == "11")
                return "ноябрь";

            if (engMont == "12")
                return "декабрь";

            return "";
        }

        //Создание пустой страницы
        void createVoidlist(ref Excel.Worksheet sheet)
        {

            Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range("A1", "DF129").Cells;
            // Производим объединение
            _excelCells1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        }

        //объеденение 
        void oneColum(ref Excel.Worksheet sheet, string start, string finish)
        {
            Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range(start, finish).Cells;
            // Производим объединение
            _excelCells1.Merge(Type.Missing);
            //_excelCells1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

        }

        //координаты
        string coordinateToString(int ch, int i)
        {
            string num = i.ToString();
            char buffSTART = '@';
            char buffEND = 'A';
            int nb = ch;
            string answer;


            if (25 < nb)
            {
                while (25 < nb)
                {
                    nb -= 26;
                    buffSTART = (char)((int)(buffSTART) + 1); 

                }
                buffEND = (char)((int)buffEND + nb);
                answer = buffSTART.ToString() + buffEND.ToString() + i.ToString();
            }
            else
            {
                buffEND = (char)((int)buffEND + nb);
                answer = buffEND.ToString() + i.ToString();
            }


            return answer;

        }

        //считывание групы
        int[] getAllLessin(string AllLessin, ref int count)
        {
            int cansel = AllLessin.Length;
            count = 0;
            string bufer = "";
            int i = 0;


            while (i != cansel )
            {
                bufer = "";
                while (i != cansel && AllLessin[i] != '-' && AllLessin[i] != '\0' && i != cansel && AllLessin[i] != ',')
                {
                    bufer += AllLessin[i].ToString();
                    i++;
                }
                int start = int.Parse(bufer);
                if (i != cansel &&  AllLessin[i] == '-')
                {
                    i++;
                    bufer = "";
                    while (i != cansel &&  AllLessin[i] != '\0' && AllLessin[i] != ',' && i != cansel)
                    {
                        bufer += AllLessin[i].ToString();
                        i++;
                    }
                    int end = int.Parse(bufer);
                    count += end - start + 1;
                    if (i != cansel)
                    {
                        if (AllLessin[i] == ',')
                            {
                                i++;
                                i++;
                            }
                    }
                }
                else
                {
                    count++;
                    if (i != cansel)
                    {
                        if (AllLessin[i] == ',')
                        {
                            i++;
                            i++;
                        }
                    }
                }
            }



            int[] answer = new int[count];
            bufer = "";
            i = 0;
            int k = 0;


            while (i != cansel &&  AllLessin[i] >= '0' && AllLessin[i] <= '9')
            {
                bufer = "";
                while (i != cansel && AllLessin[i] != '-' && AllLessin[i] != '\0' && AllLessin[i] != ',')
                {
                    bufer += AllLessin[i].ToString();
                    i++;
                }
                int start = int.Parse(bufer);

                if (i != cansel &&  AllLessin[i] == '-')
                {
                    i++;
                    bufer = "";
                    while (i != cansel && AllLessin[i] != '\0' && AllLessin[i] != ','  )
                    {
                        bufer += AllLessin[i].ToString();
                        i++;
                    }
                    int end = int.Parse(bufer);

                    for (; end >= start; k++)
                    {
                        answer[k] = start;
                        start++;
                    }

                    if (i != cansel)
                    {
                        if (AllLessin[i] == ',')
                        {
                            i++;
                           i++;
                        }
                    }

                }
                else
                {
                    answer[k] = start;
                    k++;

                    if (i != cansel)
                    {
                        if (AllLessin[i] == ',')
                        {
                             i++;
                             i++;
                        }
                    }
                    
                }
            }
            return answer;
        }

        //Русификотор дней недели
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

        //Вставка группы 3A
        //При выводе в колонку +1
        void setAllLessin (ref Excel.Worksheet sheet, ref int _ch, ref int _num, Allpeople baseItem, ref int next)
        {
            _ch += 3;
            int start = _ch;
            int buffer = next;
            //Расмотреть i при выходе из этой функции
            while ( next < baseItem._count && baseItem._all[next].kyrs == baseItem._all[buffer].kyrs)
            {
                int countAllLessin = 0;

                while (next < baseItem._count && baseItem._all[next].fac == baseItem._all[buffer].fac)
                {
                    sheet.Cells[_num + 1, _ch + 1 + countAllLessin] = baseItem._all[next].AllLessin;
                    next++;
                    countAllLessin++;
                }
                
                oneColum(ref sheet, coordinateToString(_ch , _num), coordinateToString(_ch  + countAllLessin -1, _num));
                sheet.Cells[_num, _ch + 1] = baseItem._all[buffer].fac;
                _ch += countAllLessin;

                if (baseItem._all[next].kyrs == baseItem._all[buffer].kyrs)
                {
                    buffer = next;
                }
                
                
            }
            oneColum(ref sheet, coordinateToString(start,_num - 1), coordinateToString(_ch - 1,_num -1));
            sheet.Cells[_num - 1, start+1] = baseItem._all[buffer - 1].kyrs;
        }

        //Указывай поле "Дата"
        void createOneTemplate(ref Excel.Worksheet sheet, int ch, int num)
        {

            int day = DateTime.DaysInMonth(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month);  // dateTimePicker1.Value;
            
            //A3
            oneColum(ref sheet,coordinateToString(ch+1,num), coordinateToString(ch + 2, num));//левее даты 
            sheet.Cells[num, ch + 1] = "Дата";
            oneColum(ref sheet, coordinateToString(ch, num + 1), coordinateToString(ch + 2, num+1));//группы
            sheet.Cells[num + 1, ch + 1] = "Группы";
            oneColum(ref sheet, coordinateToString(ch, num + 2), coordinateToString(ch + 2, num + 2));//уч класы
            sheet.Cells[num + 2, ch + 1] = "уч. классы";

            int dateSt = num + 3;
            int dateEn = num + 6;
            string buffer;
            for (int i = 0; i< day; i++)
            {
                oneColum(ref sheet, coordinateToString(ch, dateSt + i * 4), coordinateToString(ch , dateEn + 4*i));//даты
                buffer = "";
                buffer += (i + 1).ToString();
                buffer += ".";
                buffer += dateTimePicker1.Value.Month.ToString();
                 buffer += ".";
                buffer += dateTimePicker1.Value.Year.ToString();
                sheet.Cells[ dateSt + i * 4, ch + 1] = buffer;
                oneColum(ref sheet, coordinateToString(ch + 1, dateSt + i * 4), coordinateToString(ch + 1, dateEn + 4 * i));
                //DateTime arrayAllLessin = new DateTime(buffer);
                //ТУТ ПОЧИНИТЬ
                DateTime myDate = DateTime.Parse(buffer);
                sheet.Cells[dateSt + i * 4, ch + 2] = RussFuckingDays((string)myDate.DayOfWeek.ToString());
            }
        }

        //Считывание предметов
        public Triple[] getLess(Excel.Worksheet sheet, ref int countPenis)
        {
            int column = 5;
            int row =3;
            while (RusMont(dateTimePicker1.Value.Month.ToString()) != sheet.Cells[row, column].Text && sheet.Cells[row,column].Text != "///")
            {
                column+=4;
            }
            column++;
            int startColumn = column;
            int strartRow = row;
            row++;
            while (sheet.Cells[row, column].Text != "///")
            {
                if (sheet.Cells[row, column].Text != "")
                {
                    countPenis++;
                }
                row++;

            }
            column = startColumn;
            row = strartRow;
            Triple [] answer = new Triple[countPenis];
            countPenis = 0;
            row++;

            string buferrForCommetnst = "";
            int coiuntForComments = 0;

            while(sheet.Cells[row, column].Text != "///")
            {
                if (sheet.Cells[row, column].Text != "")
                {
                    while (sheet.Cells[row, 2].Text == sheet.Cells[row + 1, 2].Text)
                    {
                        if (sheet.Cells[row, column].Text != "")
                        {

                            answer[countPenis]._time = sheet.Cells[row, column + 1].Text;
                            answer[countPenis]._tema = sheet.Cells[row, column].Text;
                            answer[countPenis]._nameDiscipline = sheet.Cells[row, 2].Text;
                            answer[countPenis]._set = true;
                            answer[countPenis]._fack = Kastil1;
                            answer[countPenis]._coments = sheet.Cells[row, column - 1].Text;
                            buferrForCommetnst += sheet.Cells[row, 3].Text + " ";
                            countPenis++;
                            coiuntForComments++;
                        }
                        row++;
                    }
                    answer[countPenis]._time = sheet.Cells[row, column + 1].Text;
                    answer[countPenis]._tema = sheet.Cells[row, column].Text;
                    answer[countPenis]._nameDiscipline = sheet.Cells[row, 2].Text;
                    answer[countPenis]._set = true;
                    answer[countPenis]._fack = Kastil1;
                    answer[countPenis]._coments = sheet.Cells[row, column - 1].Text;
                    buferrForCommetnst += sheet.Cells[row, 3].Text;
                    coiuntForComments++;
                    for (int i = 0; i< coiuntForComments;i++)
                    {
                        answer[countPenis - i]._comentsDisciplines = buferrForCommetnst;
                    }

                    countPenis++;
                    buferrForCommetnst = "";
                    coiuntForComments = 0;
                }
                row++;
            }

            return answer;
        }

        //Подсчитывание дисциплин
        int countModulesGoodMan(string hihi)
        {
            bool itsWord = false;
            int count = 0;
            for(int i = 0; i< hihi.Length; )
            {
                while (i < hihi.Length && hihi[i] != ' ')
                {
                    itsWord = true;
                    i++;
                }
                if(itsWord)
                {
                    count++;
                    itsWord = false;
                }
                while (i < hihi.Length && hihi[i] == ' ')
                {
                    i++;
                }
            }
            return count;
        }

        //Считывание дисциплин
        string [] getModule(string hi)
        {
            hi.Trim();
            int i = 0;
            int lenght = hi.Length;
            int countModule = 0;
            string[] answer = new string[countModulesGoodMan(hi)];
            while ( i < lenght)
            {
                while (i < lenght && hi[i] != ' ')
                {
                    answer[countModule] += hi[i];
                    if (i < lenght)
                    {
                        i++;
                    }
                    
                
}
                while (i < lenght && hi[i] == ' ' )
                {
                    i++;
                }
                countModule++;
            }

            return answer;
        }

        public void getGoodMan (Excel.Worksheet shteet)
        {
            int column = 1;
            int row = 1;
            int countGroup = countGroupForTeacher(shteet);
            column = 1;            
            for (int i = 0; i < countGroup; i++) 
            {
                while (shteet.Cells[row, column].Text() != "")
                {
                    row = 3;
                    while (shteet.Cells[row,column].Text() != "")
                    {
                        teacher.push(shteet.Cells[1, column].Text(), shteet.Cells[row, column].Text(), shteet.Cells[row + 1, column].Text(), TeacherFromString(shteet.Cells[row + 2, column].Text()));
                        row += 3;
                    }
                    row = 1;
                    column++;
                }
            }
        }

        public string [] TeacherFromString (string str)
        {
            
            str = str.Trim();
            str = Regex.Replace(str, @"\s+", " ");
            int count = 1;
            for (int i = 0; i < str.Length; i++)
            {
                if (str [i] == ' ')
                {
                    count++;
                }
                
            }
            
            string[] answer = new string[count];
            count = 0;

            for (int i = 0; i < str.Length; i++)
            {
                string buffer = "";
                while (i < str.Length && str[i] != ' ')
                {
                    buffer += str[i];
                    i++;
                }
                answer[count] = buffer;
                count++;
            }
            return answer;
        }

        public int countGroupForTeacher (Excel.Worksheet shteet)
        {
            int column = 1;
            int row = 1;
            int countGroup = 0;
            while (shteet.Cells[row, column].Text != "")
            {
                countGroup++;
                column++;
            }
            return countGroup;
        }

        //Подсчет для классов2
        int countles(string str)
        {
            str.Trim();
            int count = 0;
            for (int i = 0; i< str.Length;)
            {
                while (i < str.Length && str[i] != ' ')
                {
                    i++;
                }
                count++;
                while(i < str.Length && str[i] == ' ')
                {
                    i++;
                }
            }
            return count;
        }
        //Создание массива классов
        string[] GetArrrayLesss(string str, int countS)
        {
            string[] array = new string[countS];
            str.Trim();
            int count = 0;
            for (int i = 0; i < str.Length;)
            {
                while (i < str.Length && str[i] != ' ')
                {
                    array[count] += str[i];
                    i++;
                }
                count++;
                while (i < str.Length && str[i] == ' ')
                {
                    i++;
                }
            }
            return array;
        }
        
        //Подсчет для классов1
        int countModules(Excel.Worksheet shteet)
        {
            int row = 1;
            int column = 1;
            int count = 0;
            while (shteet.Cells[row,column].Text() != "")
            {
                count++;
                row++;
            }
            return count/2;
        }

        bool itsNotChange(string str)
        {
            if (str == "*")
                return true;

            return false;
        }
        //Загрузка крепления классов
        void PKP(Excel.Worksheet shteet)
        {
            int row = 1, column = 1;
            clas.setCountDGD(countModules(shteet));
            while(shteet.Cells[row,column].Text() != "")
            {
                clas.setCountModule(countles(shteet.Cells[row + 1, column].Text()) );
                clas.setModule(shteet.Cells[row, column].Text(), GetArrrayLesss(shteet.Cells[row + 1, column].Text(), countles(shteet.Cells[row + 1, column].Text())), itsNotChange(shteet.Cells[row, column + 1].Text()));
                row += 2;
            }
        }

        //Подсчет сюда не лезь
        int countGP(Excel.Worksheet shteet)
        {
            int count = 0, row = 2, column = 2;
            while (shteet.Cells[row,column].Text() != "")
            {
                while (shteet.Cells[row + 1, column].Text() != "")
                {
                    count++;
                    row++;
                }
                row = 2;
                column++;
            }

            return count;
        }
        
        //Подсчет сюда не лезь
        void getGP(Excel.Worksheet shteet)
        {

            AllLessinRooms.setCountArray(countGP(shteet));
            int row = 2, column = 2;
            while (shteet.Cells[row, column].Text() != "")
            {
                int name = 2;
                while (shteet.Cells[row + 1, column].Text() != "")
                {
                    AllLessinRooms.pushGP(shteet.Cells[row + 1, column].Text(), shteet.Cells[name, column].Text());
                    row++;
                }
                row = name;
                column++;
            }

        }

        //Количество столбцов в новой форме
        int CountColumnForNewForms(int countAllLessin, int countUploadFile)
        {
            int answer = 3 * 5 + countAllLessin * 2;

            return answer;
        }

        //Количество строчек
        int CountRowForNewForms ()
        {
            int daysInMonth = DateTime.DaysInMonth(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month);

            return daysInMonth*4 + 4;
        }

        string NaxuiLishnueProbels(string e)
        {
            
            string answer ="";
            for (int i = 0; i< e.Length-1;i++)
            {
                if (e.Length >1 && (e[i] != ' ' || e[i+1] != ' '))
                {
                    answer += e[i];
                }
                
            }
            if (e.Length >1)
            {
                answer += e[e.Length - 1];
            }
            
            return answer;
        }
        //Параметры соеднинения
        public string get_cs()
        {
            return "Data Source="+ Properties.Settings.Default.PathSqlServer + "; Initial Catalog =DarkLight ; User ID = sa; Password = 123456";
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 form = new Form3();
            form.ShowDialog();
        }
    }
    public struct Triple
    {
        public string _comentsDisciplines;
        public string _coments;
        public string _time;
        public string _tema;
        public string _nameDiscipline;
        public int _fack;
        public bool _set ;

    }
    public struct AllLessinRomms
    {
        public string rooms;
        public string AllLessin;
    }
    public struct CHEA
    {
        public bool _nChange;
        public string _nameModule;
        public string[] _rooms;
    }
    public struct Module
    {
        public string _nameModule;
        public string[] _elemetntsModule;
        public int[] _numberAllLessin;
    }
    public class Allpeople
    {
        public XoordinateKG [] _all;
        public int _count;
        public string [][][] _allNewForNewForms;
        private int _countForBigBoss;
        private int _countForGetOneAllLessin;

        public Allpeople()
        {
            _all = new XoordinateKG[0];
            _count = 0;
            _countForBigBoss = 0;
            _countForGetOneAllLessin = 0;
        }

        ~Allpeople()
        {
            
        }

        public void pushElements(string fac, int AllLessin , string kyrs , string clas = "" )
        {
            XoordinateKG[] buffer = new XoordinateKG[_count + 1];
            for(int i = 0; i < _count; i++)
            {
                buffer[i] = _all[i];
            }

            buffer[_count].AllLessin = AllLessin;
            buffer[_count].fac = fac;
            buffer[_count].clas = clas;
            buffer[_count].kyrs = kyrs;
            _all = buffer;
            _count++;
        }

        // Количетво групп в курсе, при последущем вызове обновляет курс
        public int BigBoss()
        {
            int count = 1;


            while (_countForBigBoss < _count - 1  &&_all[_countForBigBoss].kyrs == _all[_countForBigBoss + 1].kyrs) 
            {
                _countForBigBoss++;
                count++;
            }

            if (_count == _countForBigBoss)
            {
                _countForBigBoss = 0;
            }
            _countForBigBoss++;
            return count;

        }

        public void GetOneAllLessin (DataGridView dategridView, int row, int column, int space, int allRow, CellsTable[,] arrayTable)
        {
            int count = 0;
             while (_count > _countForGetOneAllLessin + 1 && _all[_countForGetOneAllLessin].kyrs == _all[_countForGetOneAllLessin + 1].kyrs)
            {
                int pipidastr = _countForGetOneAllLessin;
                while (_count > _countForGetOneAllLessin+1 && _all[_countForGetOneAllLessin].fac == _all[_countForGetOneAllLessin + 1].fac && _all[_countForGetOneAllLessin].kyrs == _all[_countForGetOneAllLessin + 1].kyrs)
                {
                    dategridView[column + count, row].Value = _all[_countForGetOneAllLessin].AllLessin;
                    _countForGetOneAllLessin++;
                    count++;
                }
                
                dategridView[column + count, row].Value = _all[_countForGetOneAllLessin].AllLessin;
                count++;

                while (_count > pipidastr+1 && _all[pipidastr].fac == _all[pipidastr + 1].fac && _all[pipidastr].kyrs == _all[pipidastr + 1].kyrs)
                {
                    dategridView[column + count, row].Value = _all[pipidastr].AllLessin;
                    for (int i = 0; i < allRow; i++)
                    {
                        arrayTable[i, column + count]._down = true;
                    }
                    pipidastr++;
                    count++;
                    
                }
                dategridView[column + count, row].Value = _all[pipidastr].AllLessin;
                for (int i = 0; i < allRow; i++)
                {
                    arrayTable[i, column + count]._down = true;
                }
                count++;

                if (_count > _countForGetOneAllLessin + 1 && _all[_countForGetOneAllLessin].kyrs != _all[_countForGetOneAllLessin + 1].kyrs)
                {
                    _countForGetOneAllLessin++;
                    break;
                }
                else
                {
                    if (_count != _countForGetOneAllLessin + 1)
                    _countForGetOneAllLessin++;
                }
            }

            if (_count == _countForGetOneAllLessin)
            {
                _countForGetOneAllLessin = 0;
            }
        }


    }
    public struct XoordinateKG
        {
            public string fac;
            public int AllLessin;
            public string clas;
            public string kyrs;
        }
    public class AllLessin
    {
        public Triple[] _arrayLesson;
        public string _nameAllLessin;
        private int _count = 0;

        public AllLessin(string nameAllLessin, int count)
        {
            _arrayLesson = new Triple[count];
            _nameAllLessin = nameAllLessin;
            _count = count;
        }

        ~AllLessin()
        {

        }
        
        public void pushtriple (Triple []nameLesson)
        {
            _arrayLesson = nameLesson;
        }
        public void pushOneTriple (Triple triple)
        {
            Triple [] nowHow = new Triple[_count + 1];
            for(int i =0; i< _arrayLesson.Length; i++)
            {
                nowHow[i] = _arrayLesson[i];
            }
            nowHow[_count] = triple;
            _count++;
            _arrayLesson = nowHow;
        }

    }
    public struct GroupeTeacher
    {
        public string _numberGroupe;
        public string _nameDisiplines;
        public string _tems;
        public string[] _teachers;
    }
    public class Teacher
    {
        public GroupeTeacher[] _array;
        private int _count;

        public Teacher()
        {
            _count = 0;
            _array = new GroupeTeacher[_count];
        }

        public void push(string numberGroupe, string nameDisiplines, string tems, string [] teachers)
        {
            GroupeTeacher[] answer = new GroupeTeacher[_count + 1];
            for(int i = 0; i < _count; i++)
            {
                answer[i] = _array[i];
            }
            answer[_count]._nameDisiplines = nameDisiplines;
            answer[_count]._numberGroupe = numberGroupe;
            answer[_count]._tems = tems;
            answer[_count]._teachers = teachers;
            _array = answer;
            _count++;
        }

        public int getCount() { return _count; }
    }
    public class KMA
    {
        protected CHEA [] _DGD;
        protected int _number;
        public KMA()
        {
            _number = 0;
        }

        public void setCountDGD (int count)
        {
            _DGD = new CHEA[count];
        }

        public void setCountModule(int count)
        {
            _DGD[_number]._rooms = new string[count];
        }

        public void setModule(string name,string [] array, bool change)
        {
            _DGD[_number]._nameModule = name;
            _DGD[_number]._nChange = change;
            _DGD[_number]._rooms = array;
            _number++;
        }
        
        ~KMA()
        {

        }

    }
    public class AllLessinAndRooms
    {
        public AllLessinRomms[] _array;
        public int _number;

        public AllLessinAndRooms()
        {
            _number = 0;
            _array = new AllLessinRomms[0];
        }

        public void pushGPSQL(string room, string name)
        {
            AllLessinRomms [] buffer = new AllLessinRomms[_number + 1];
            for(int i =0; i < _array.Length; i++)
            {
                buffer[i] = _array[i];
            }
            buffer[_number].rooms = room;
            buffer[_number].AllLessin = name;
            _number++;
            _array = buffer;

        }

        public void pushGP(string room, string AllLessin)
        {
            _array[_number].AllLessin = AllLessin;
            _array[_number].rooms = room;
            _number++;
        }

        public void setCountArray(int count)
        {
            _array = new AllLessinRomms[count];
        }
    }
}