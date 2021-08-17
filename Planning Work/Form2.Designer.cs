namespace Planning_Work
{
    partial class Form2
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.button5 = new System.Windows.Forms.Button();
            this.colorDialog2 = new System.Windows.Forms.ColorDialog();
            this.button6 = new System.Windows.Forms.Button();
            this.colorDialog3 = new System.Windows.Forms.ColorDialog();
            this.button7 = new System.Windows.Forms.Button();
            this.colorDialog4 = new System.Windows.Forms.ColorDialog();
            this.button8 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1670, 1080);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);
            this.dataGridView1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dataGridView1_CurrentCellDirtyStateChanged);
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError);
            this.dataGridView1.DragDrop += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragDrop);
            this.dataGridView1.DragOver += new System.Windows.Forms.DragEventHandler(this.dataGridView1_DragOver);
            this.dataGridView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDown);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1676, 970);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(236, 46);
            this.button1.TabIndex = 1;
            this.button1.Text = "Размер ячеек и текста ";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(1676, 918);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(236, 46);
            this.button4.TabIndex = 13;
            this.button4.Text = "Сохранить ";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(1676, 882);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(236, 30);
            this.button5.TabIndex = 14;
            this.button5.Text = "Задать цвет совпадения преподавателей\r\n";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(1676, 846);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(236, 30);
            this.button6.TabIndex = 15;
            this.button6.Text = "Задать цвет совпадение аудиторий\r\n";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(1676, 774);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(236, 30);
            this.button7.TabIndex = 16;
            this.button7.Text = "Задать цвет данному предмету";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(1676, 810);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(236, 30);
            this.button8.TabIndex = 17;
            this.button8.Text = "Задать цвет нарушения логики";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(1676, 751);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(156, 17);
            this.checkBox1.TabIndex = 18;
            this.checkBox1.Text = "Режим заполнения ячеек";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(1752, 122);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 7;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1766, 135);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 9;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(1768, 148);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(0, 13);
            this.label8.TabIndex = 11;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1676, 1022);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(236, 46);
            this.button2.TabIndex = 21;
            this.button2.Text = "Выход";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button9
            // 
            this.button9.BackColor = System.Drawing.Color.White;
            this.button9.BackgroundImage = global::Planning_Work.Properties.Resources.cut_105155;
            this.button9.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button9.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.button9.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button9.Location = new System.Drawing.Point(1676, 12);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(90, 84);
            this.button9.TabIndex = 19;
            this.button9.UseVisualStyleBackColor = false;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(1676, 239);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(147, 13);
            this.label1.TabIndex = 22;
            this.label1.Text = "Коментарий к данной паре:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1676, 252);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 23;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(1920, 1080);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form2";
            this.Text = "Расписание";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Form2_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ColorDialog colorDialog2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.ColorDialog colorDialog3;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.ColorDialog colorDialog4;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}