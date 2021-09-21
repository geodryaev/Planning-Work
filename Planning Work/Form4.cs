using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Planning_Work
{
    public partial class Form4 : Form
    {
        int pozitionX = 170, pozitionY = 210;
        bool chaekFullUp;
        ComboBox[] arr;
        int count_box;
        CellsTable[,] cellsTables;
        DataGridViewCellEventArgs ob;
        DataGridView dataGridView;
        string saveTem, saveDisiplines, saveNameGroupe;
        Teacher saveTeacher;
        public Form4(string nameDisciplines, string tem, ref CellsTable[,] _arrayTable, ref object sender, ref DataGridViewCellEventArgs e, AllLessinAndRooms clas, Teacher teacher, string nameGroupe, ref DataGridView dataGridView1)
        {
            InitializeComponent();
            cellsTables = _arrayTable;
            dataGridView = dataGridView1;
            ob = e;

            saveTem = tem;
            saveTeacher = teacher;
            saveDisiplines = nameDisciplines;
            saveNameGroupe = nameGroupe;

            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;//Класс
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;//Имя предмета
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;//Тип
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;//Имя преподавателя
            arr = new ComboBox[4];
            count_box = 1;
            arr[0] = comboBox4;


            comboBox2.Items.Add (nameDisciplines);
            comboBox3.Items.Add(tem);
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            
            for (int i = 0;  i< clas._array.Length; i++)
            {
                if (clas._array[i].AllLessin == nameDisciplines)
                    comboBox1.Items.Add(clas._array[i].rooms);
            }



            if (searchTem(teacher, tem, nameDisciplines, nameGroupe) == -1)
            {
                for (int i = 0; i < teacher._array.Length; i++)
                {
                    if (teacher._array[i]._nameDisiplines == nameDisciplines && teacher._array[i]._numberGroupe == nameGroupe)
                    {
                        for (int j = 0; j< teacher._array[i]._teachers.Length; j++)
                        {
                            comboBox4.Items.Add(teacher._array[i]._teachers[j]);
                        }
                    }

                }
            }
            else
            {
                int count = searchTem(teacher, tem, nameDisciplines, nameGroupe);
                for (int i = 0; i < teacher._array[count]._teachers.Length; i++)
                {
                    comboBox4.Items.Add(teacher._array[count]._teachers[i]);
                }
            }
            
            
        }

        public Form4(string nameDisciplines, string tem, string [] teachh, string roomss, ref CellsTable[,] _arrayTable, ref object sender, ref DataGridViewCellEventArgs e, AllLessinAndRooms clas, Teacher teacher, string nameGroupe, ref DataGridView dataGridView1)
        {
            InitializeComponent();
            cellsTables = _arrayTable;
            dataGridView = dataGridView1;
            ob = e;
            saveTem = tem;
            saveTeacher = teacher;
            saveDisiplines = nameDisciplines;
            saveNameGroupe = nameGroupe;

            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;//Класс
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;//Имя предмета
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;//Тип
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;//Имя преподавателя

            comboBox1.Items.Add(roomss);
            comboBox2.Items.Add(nameDisciplines);
            comboBox3.Items.Add(tem);
            Controls.Remove(comboBox4);
            
            arr = new ComboBox[4];
            arr[0] = comboBox4;
            count_box = 0;
            pozitionY -= 45;
            int countTeachForArray = 0;
            foreach (var item in teachh)
            {
                arr[count_box] = new ComboBox() { Location = new Point(pozitionX, pozitionY + 45), Width = 278, Height = 21, DropDownStyle = ComboBoxStyle.DropDownList, Name = Convert.ToString(count_box) };
                Controls.Add(arr[count_box]);
                arr[count_box].Items.Add(teachh[countTeachForArray]);
                arr[count_box].SelectedIndex = 0;
                countTeachForArray++;

                if (searchTem(saveTeacher, saveTem, saveDisiplines, saveNameGroupe) == -1)
                {
                    for (int i = 0; i < saveTeacher._array.Length; i++)
                    {
                        if (saveTeacher._array[i]._nameDisiplines == saveDisiplines && saveTeacher._array[i]._numberGroupe == saveNameGroupe)
                        {
                            for (int j = 0; j < saveTeacher._array[i]._teachers.Length; j++)
                            {
                                arr[count_box].Items.Add(saveTeacher._array[i]._teachers[j]);
                            }
                        }

                    }
                }
                else
                {
                    int count = searchTem(saveTeacher, saveTem, saveDisiplines, saveNameGroupe);
                    for (int i = 0; i < saveTeacher._array[count]._teachers.Length; i++)
                    {
                        arr[count_box].Items.Add(saveTeacher._array[count]._teachers[i]);
                    }
                }
                pozitionY += 45;
                count_box++;
            }

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;

            for (int i = 0; i < clas._array.Length; i++)
            {
                if (clas._array[i].AllLessin == nameDisciplines)
                    comboBox1.Items.Add(clas._array[i].rooms);
            }

            if (searchTem(teacher, tem, nameDisciplines, nameGroupe) == -1)
            {
                for (int i = 0; i < teacher._array.Length; i++)
                {
                    if (teacher._array[i]._nameDisiplines == nameDisciplines && teacher._array[i]._numberGroupe == nameGroupe)
                    {
                        for (int j = 0; j < teacher._array[i]._teachers.Length; j++)
                        {
                            comboBox4.Items.Add(teacher._array[i]._teachers[j]);
                        }
                    }
                }
            }
            else
            {
                int count = searchTem(teacher, tem, nameDisciplines, nameGroupe);
                for (int i = 0; i < teacher._array[count]._teachers.Length; i++)
                {
                    comboBox4.Items.Add(teacher._array[count]._teachers[i]);
                }
            }
        }

        //OK
        private void button1_Click(object sender, EventArgs e)
        {
            chaekFullUp = true;
            if (comboBox1.Text != "" && comboBox2.Text != "" && comboBox3.Text != "" )
            {
                foreach (var item in arr)
                {
                    if (item != null && item.Text == "")
                    {
                        chaekFullUp = false;
                    }
                }

                if (chaekFullUp)
                {
                    int up = countUp(), down = countDown();
                    set(up, down);
                    Close();
                }
                else
                {
                    MessageBox.Show("Пожалуйста заполните все поля ввода");
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста заполните все поля ввода");
            }
            
        }


        public int GetDistance
        {
            get { return countUp() + countDown(); }
        }
        //OTMEHA
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы хотите выйти не записав данные, продолжить ?", "Внимание", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Close();
            }
            else if (dialogResult == DialogResult.No)
            {
                
            }
            
        }

        public int searchTem (Teacher teach, string tem, string nameDisicp, string nameGroupe)
        {
            for(int i = 0; i < teach._array.Length; i ++)
            {
                if (teach._array[i]._tems == tem && teach._array[i]._nameDisiplines == nameDisicp && teach._array[i]._numberGroupe == nameGroupe)
                    return i;
            }

            return -1;
        }

        public int countUp()
        {
            int answer = 0;
            int row = ob.RowIndex, column = ob.ColumnIndex;

            for (int i = row - 3; i < row; i++)
                if (cellsTables[row, column]._disiplines == cellsTables[i, column]._disiplines && cellsTables[row, column]._tema == cellsTables[i, column]._tema)
                    answer++;

            return answer;
        }

        public int countDown()
        {
            int answer = 0;
            int row = ob.RowIndex, column = ob.ColumnIndex;
            for (int i = row + 3; i > row; i--)
                if (cellsTables[row, column]._disiplines == cellsTables[i, column]._disiplines && cellsTables[row, column]._tema == cellsTables[i, column]._tema)
                    answer++;

            return answer;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(count_box - 1 >= 1)
            {
                Controls.Remove(arr[count_box - 1]);
                count_box--;
                pozitionY -= 45;
            }
        }

        public void set(int up, int down)
        {
            string boba = "";
            string[] paxaSexs = new string [count_box];
            int row = ob.RowIndex, column = ob.ColumnIndex;
            for (int i = row - up; i <= row+down;i++ )
            {
                boba = "";
                paxaSexs = new string[count_box];
                cellsTables[i, column]._disiplines = comboBox2.Text;
                cellsTables[i, column]._rooms = comboBox1.Text;
                for (int j =0; j < count_box;j++)
                {
                    paxaSexs[j] = arr[j].Text;
                }
                cellsTables[i, column]._teacher = paxaSexs;
                cellsTables[i, column]._tema = comboBox3.Text;
                foreach (var item in paxaSexs)
                {
                    boba += item;
                    boba += " ";
                }
                dataGridView[column, i].Value = comboBox2.Text + " " + comboBox3.Text + " " + comboBox1.Text + " " + boba;
                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (count_box <= 3)
            {
                arr[count_box] = new ComboBox() { Location = new Point(pozitionX, pozitionY + 45), Width = 278, Height = 21, DropDownStyle = ComboBoxStyle.DropDownList, Name = Convert.ToString(count_box) };
                Controls.Add(arr[count_box]);

                if (searchTem(saveTeacher, saveTem, saveDisiplines, saveNameGroupe) == -1)
                {
                    for (int i = 0; i < saveTeacher._array.Length; i++)
                    {
                        if (saveTeacher._array[i]._nameDisiplines == saveDisiplines && saveTeacher._array[i]._numberGroupe == saveNameGroupe)
                        {
                            for (int j = 0; j < saveTeacher._array[i]._teachers.Length; j++)
                            {
                                arr[count_box].Items.Add(saveTeacher._array[i]._teachers[j]);
                            }
                        }

                    }
                }
                else
                {
                    int count = searchTem(saveTeacher, saveTem, saveDisiplines, saveNameGroupe);
                    for (int i = 0; i < saveTeacher._array[count]._teachers.Length; i++)
                    {
                        arr[count_box].Items.Add(saveTeacher._array[count]._teachers[i]);
                    }
                }
                count_box++;
                pozitionY += 45;

            }
        }
    }
}
