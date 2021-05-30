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
        CellsTable[,] cellsTables;
        DataGridViewCellEventArgs ob;
        DataGridView dataGridView;
        public Form4(string nameDisciplines, string tem, ref CellsTable[,] _arrayTable, ref object sender, ref DataGridViewCellEventArgs e, AllLessinAndRooms clas, Teacher teacher, string nameGroupe, ref DataGridView dataGridView1)
        {
            InitializeComponent();
            cellsTables = _arrayTable;
            dataGridView = dataGridView1;
            ob = e;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;//Класс
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;//Имя предмета
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;//Тип
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;//Имя преподавателя

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

        public Form4(string nameDisciplines, string tem, string teachh, string roomss, ref CellsTable[,] _arrayTable, ref object sender, ref DataGridViewCellEventArgs e, AllLessinAndRooms clas, Teacher teacher, string nameGroupe, ref DataGridView dataGridView1)
        {
            InitializeComponent();
            cellsTables = _arrayTable;
            dataGridView = dataGridView1;
            ob = e;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;//Класс
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;//Имя предмета
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;//Тип
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;//Имя преподавателя

            comboBox1.Items.Add(roomss);
            comboBox2.Items.Add(nameDisciplines);
            comboBox3.Items.Add(tem);
            comboBox4.Items.Add(teachh);
            
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;

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
            if (comboBox1.Text != "" && comboBox2.Text != "" && comboBox3.Text != "" && comboBox4.Text != "" )
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

        public void set(int up, int down)
        {
            int row = ob.RowIndex, column = ob.ColumnIndex;
            for (int i = row - up; i <= row+down;i++ )
            {
                cellsTables[i, column]._disiplines = comboBox2.Text;
                cellsTables[i, column]._rooms = comboBox1.Text;
                cellsTables[i, column]._teacher = comboBox4.Text;
                cellsTables[i, column]._tema = comboBox3.Text;
                dataGridView[column, i].Value = comboBox2.Text + " " + comboBox3.Text + " " + comboBox1.Text + " " + comboBox4.Text;
            }
        }
    }
}
