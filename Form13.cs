using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Veles
{
    public partial class Form13 : Form
    {
        public Form13()
        {
            InitializeComponent();
            dataGridView1.DataError += new DataGridViewDataErrorEventHandler(dataGridView1_DataError); 
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
        } //Обход ошибки о нулевом значении ячейки. Необходимо для combobox'ов.

        private void Form13_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Чип". При необходимости она может быть перемещена или удалена.
            this.чипTableAdapter.Fill(this.заправка1DataSet.Чип);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Лезвие_очистки". При необходимости она может быть перемещена или удалена.
            this.лезвие_очисткиTableAdapter.Fill(this.заправка1DataSet.Лезвие_очистки);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.PCR". При необходимости она может быть перемещена или удалена.
            this.pCRTableAdapter.Fill(this.заправка1DataSet.PCR);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Ракель". При необходимости она может быть перемещена или удалена.
            this.ракельTableAdapter.Fill(this.заправка1DataSet.Ракель);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Магнитный_вал". При необходимости она может быть перемещена или удалена.
            this.магнитный_валTableAdapter.Fill(this.заправка1DataSet.Магнитный_вал);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Барабан". При необходимости она может быть перемещена или удалена.
            this.барабанTableAdapter.Fill(this.заправка1DataSet.Барабан);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Тонер". При необходимости она может быть перемещена или удалена.
            this.тонерTableAdapter.Fill(this.заправка1DataSet.Тонер);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Производитель". При необходимости она может быть перемещена или удалена.
            this.производительTableAdapter.Fill(this.заправка1DataSet.Производитель);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Картридж". При необходимости она может быть перемещена или удалена.
            this.картриджTableAdapter.Fill(this.заправка1DataSet.Картридж);

        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Производитель (Производитель) VALUES (@Proizvod);";
            var Proizvod = comboBox1.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("Proizvod", Proizvod);
            if (Proizvod == "")
            {
                MessageBox.Show("Укажите название производителя!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.производительTableAdapter.Fill(this.заправка1DataSet.Производитель);
            }
        } //Добавление в БД нового производителя

        private void button2_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Картридж ([Название картриджа],Производитель, Тонер," +
                                        "Барабан, Ракель, PCR, [Магнитный вал],[Лезвие очистки], Чип)" +
                                        "VALUES (@NameKar, @Proizv, @Toner, @Baraban, @Rakel, @PCR, @MagVal, @LezOch, @Chip);";
            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);


            string NameKar = comboBox2.Text;
            insertCommand.Parameters.AddWithValue("NameKar", comboBox2.Text);
            if (comboBox1.Text != "")
            {
                int Proizv = (int)comboBox1.SelectedValue;
                insertCommand.Parameters.AddWithValue("Proizv", Proizv);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("Proizv", DBNull.Value);
            }
            if (comboBox3.Text != "")
            {
                int Toner = (int)comboBox3.SelectedValue;
                insertCommand.Parameters.AddWithValue("Toner", Toner);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("Toner", DBNull.Value);
            }
            if (comboBox4.Text != "")
            {
                int Baraban = (int)comboBox4.SelectedValue;
                insertCommand.Parameters.AddWithValue("Baraban", Baraban);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("Baraban", DBNull.Value);
            }
            if (comboBox5.Text != "")
            {
                int Rakel = (int)comboBox5.SelectedValue;
                insertCommand.Parameters.AddWithValue("Rakel", Rakel);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("Rakel", DBNull.Value);
            }
            if (comboBox8.Text != "")
            {
                int PCR = (int)comboBox8.SelectedValue;
                insertCommand.Parameters.AddWithValue("PCR", PCR);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("PCR", DBNull.Value);
            }
            if (comboBox6.Text != "")
            {
                int MagVal = (int)comboBox6.SelectedValue;
                insertCommand.Parameters.AddWithValue("MagVal", MagVal);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("MagVal", DBNull.Value);
            }
            if (comboBox7.Text != "")
            {
                int LezOch = (int)comboBox7.SelectedValue;
                insertCommand.Parameters.AddWithValue("LezOch", LezOch);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("LezOch", DBNull.Value);
            }
            if (comboBox9.Text != "")
            {
                int Chip = (int)comboBox9.SelectedValue;
                insertCommand.Parameters.AddWithValue("Chip", Chip);
            }
            else
            {
                insertCommand.Parameters.AddWithValue("Chip", DBNull.Value);
            }

            if (comboBox2.Text != "")
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.картриджTableAdapter.Fill(this.заправка1DataSet.Картридж);
            }
            else
            {
                MessageBox.Show("Введите название картриджа!");
            }
        } //Добавление в БД нового картриджа

        private void button3_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Тонер (Тонер, Цена) VALUES (@Toner, @Price);";

            string Toner = comboBox3.Text;
            string Price = comboBox10.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("Toner", Toner);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (Toner == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену тонера!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.тонерTableAdapter.Fill(this.заправка1DataSet.Тонер);
            }
        } //Добавление в БД нового тонера

        private void button4_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Барабан (Барабан, Цена) VALUES (@Baraban, @Price);";

            string Baraban = comboBox4.Text;
            string Price = comboBox11.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("Baraban", Baraban);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (Baraban == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену барабана!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.барабанTableAdapter.Fill(this.заправка1DataSet.Барабан);
            }
        } //Добавление в БД нового барабана

        private void button6_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO [Магнитный вал] ([Магнитный вал], Цена) VALUES (@MagnVal, @Price);";

            string MagnVal = comboBox6.Text;
            string Price = comboBox12.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("MagnVal", MagnVal);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (MagnVal == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену магнитного вала!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.магнитный_валTableAdapter.Fill(this.заправка1DataSet.Магнитный_вал);
            }
        } //Добавление в БД нового магнитного вала

        private void button5_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Ракель (Ракель, Цена) VALUES (@Rakel, @Price);";

            string Rakel = comboBox5.Text;
            string Price = comboBox14.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("Rakel", Rakel);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (Rakel == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену ракеля!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.ракельTableAdapter.Fill(this.заправка1DataSet.Ракель);
            }
        } //Добавление в БД нового ракеля

        private void button8_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO PCR (PCR, Цена) VALUES (@PCR, @Price);";

            string PCR = comboBox8.Text;
            string Price = comboBox6.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("PCR", PCR);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (PCR == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену PCR!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.pCRTableAdapter.Fill(this.заправка1DataSet.PCR);
            }
        } //Добавление в БД нового PCR

        private void button7_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO [Лезвие очистки] ([Лезвие очистки], Цена) VALUES (@LezvOchist, @Price);";

            string LezvOchist = comboBox7.Text;
            string Price = comboBox15.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("LezvOchist", LezvOchist);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (LezvOchist == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену лезвия очистки!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.лезвие_очисткиTableAdapter.Fill(this.заправка1DataSet.Лезвие_очистки);
            }
        } //Добавление в БД нового лезвия очистки

        private void button9_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "INSERT INTO Чип (Чип, Цена) VALUES (@Chip, @Price);";

            string Chip = comboBox9.Text;
            string Price = comboBox16.Text;

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            insertCommand.Parameters.AddWithValue("Chip", Chip);
            insertCommand.Parameters.AddWithValue("Price", Price);
            if (Chip == "" || Price == "")
            {
                MessageBox.Show("Укажите название и цену чипа!");
            }
            else
            {
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно добавлена!");
                this.чипTableAdapter.Fill(this.заправка1DataSet.Чип);
            }
        } //Добавление в БД нового чипа

        private void button10_Click(object sender, EventArgs e)
        {
            {
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Картридж " +
                    "SET Производитель = @Proizv, " +
                    "Тонер = @Toner, Барабан = @Baraban, Ракель = @Rakel, PCR = @PCR, " +
                    "[Магнитный вал] = @MagVal, [Лезвие очистки] = @LezOch, Чип = @Chip " +
                    "WHERE Код = @KodKar";

                OleDbConnection conn = new OleDbConnection(ConnectionString);
                OleDbCommand insertCommand = new OleDbCommand(commandText, conn);



                if (comboBox1.Text != "")
                {
                    int Proizv = (int)comboBox1.SelectedValue;
                    insertCommand.Parameters.AddWithValue("Proizv", Proizv);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("Proizv", DBNull.Value);
                }
                if (comboBox3.Text != "")
                {
                    int Toner = (int)comboBox3.SelectedValue;
                    insertCommand.Parameters.AddWithValue("Toner", Toner);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("Toner", DBNull.Value);
                }
                if (comboBox4.Text != "")
                {
                    int Baraban = (int)comboBox4.SelectedValue;
                    insertCommand.Parameters.AddWithValue("Baraban", Baraban);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("Baraban", DBNull.Value);
                }
                if (comboBox5.Text != "")
                {
                    int Rakel = (int)comboBox5.SelectedValue;
                    insertCommand.Parameters.AddWithValue("Rakel", Rakel);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("Rakel", DBNull.Value);
                }
                if (comboBox8.Text != "")
                {
                    int PCR = (int)comboBox8.SelectedValue;
                    insertCommand.Parameters.AddWithValue("PCR", PCR);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("PCR", DBNull.Value);
                }
                if (comboBox6.Text != "")
                {
                    int MagVal = (int)comboBox6.SelectedValue;
                    insertCommand.Parameters.AddWithValue("MagVal", MagVal);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("MagVal", DBNull.Value);
                }
                if (comboBox7.Text != "")
                {
                    int LezOch = (int)comboBox7.SelectedValue;
                    insertCommand.Parameters.AddWithValue("LezOch", LezOch);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("LezOch", DBNull.Value);
                }
                if (comboBox9.Text != "")
                {
                    int Chip = (int)comboBox9.SelectedValue;
                    insertCommand.Parameters.AddWithValue("Chip", Chip);
                }
                else
                {
                    insertCommand.Parameters.AddWithValue("Chip", DBNull.Value);
                }

                int KodKar = (int)comboBox2.SelectedValue;
                insertCommand.Parameters.AddWithValue("KodKar", KodKar);

                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Запись успешно изменена!");
                this.картриджTableAdapter.Fill(this.заправка1DataSet.Картридж);
            }
        }//Редактирование нужной строки


    }
}
