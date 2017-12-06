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
using Microsoft.Office;

namespace Veles
{
    public partial class Form14 : Form
    {
        public Form14()
        {
            InitializeComponent();
        }

        private void Form14_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Таблица_для_Тендера". При необходимости она может быть перемещена или удалена.
            this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "заправка1DataSet.Тендер". При необходимости она может быть перемещена или удалена.
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);

            таблица_для_ТендераTableAdapter.Update(заправка1DataSet);
            this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);

            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "DELETE * FROM Тендер;";

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            conn.Open();
            insertCommand.ExecuteNonQuery();
            conn.Close();
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);

        }
        public void TableCheker()
        {
            if (checkBox2.Checked == false) //тонер
            {
                dataGridView2.Columns["тонерШтуки"].Visible = false;
                dataGridView2.Columns["тонерМасса"].Visible = false;
                dataGridView2.Columns["тонерЦена"].Visible = false;
                dataGridView2.Columns["тонерСтоимость"].Visible = false;
                dataGridView2.Columns["тонерНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["тонерШтуки"].Visible = true;
                dataGridView2.Columns["тонерМасса"].Visible = true;
                dataGridView2.Columns["тонерЦена"].Visible = true;
                dataGridView2.Columns["тонерСтоимость"].Visible = true;
                dataGridView2.Columns["тонерНДС"].Visible = true;
            }
            if (checkBox4.Checked == false) //Барабан
            {
                dataGridView2.Columns["барабанШтуки"].Visible = false;
                dataGridView2.Columns["барабанЦена"].Visible = false;
                dataGridView2.Columns["барабанСтоимость"].Visible = false;
                dataGridView2.Columns["барабанНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["барабанШтуки"].Visible = true;
                dataGridView2.Columns["барабанЦена"].Visible = true;
                dataGridView2.Columns["барабанСтоимость"].Visible = true;
                dataGridView2.Columns["барабанНДС"].Visible = true;
            }
            if (checkBox5.Checked == false) //Ракель
            {
                dataGridView2.Columns["ракельШтуки"].Visible = false;
                dataGridView2.Columns["ракельЦена"].Visible = false;
                dataGridView2.Columns["ракельСтоимость"].Visible = false;
                dataGridView2.Columns["ракельНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["ракельШтуки"].Visible = true;
                dataGridView2.Columns["ракельЦена"].Visible = true;
                dataGridView2.Columns["ракельСтоимость"].Visible = true;
                dataGridView2.Columns["ракельНДС"].Visible = true;
            }
            if (checkBox1.Checked == false)//PCR
            {
                dataGridView2.Columns["pCRШтуки"].Visible = false;
                dataGridView2.Columns["pCRЦена"].Visible = false;
                dataGridView2.Columns["pCRСтоимость"].Visible = false;
                dataGridView2.Columns["pCRНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["pCRШтуки"].Visible = true;
                dataGridView2.Columns["pCRЦена"].Visible = true;
                dataGridView2.Columns["pCRСтоимость"].Visible = true;
                dataGridView2.Columns["pCRНДС"].Visible = true;
            }
            if (checkBox6.Checked == false)//Магнитный вал
            {
                dataGridView2.Columns["магнитныйВалШтуки"].Visible = false;
                dataGridView2.Columns["магнитныйВалЦена"].Visible = false;
                dataGridView2.Columns["магнитныйВалСтоимость"].Visible = false;
                dataGridView2.Columns["магнитныйВалНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["магнитныйВалШтуки"].Visible = true;
                dataGridView2.Columns["магнитныйВалЦена"].Visible = true;
                dataGridView2.Columns["магнитныйВалСтоимость"].Visible = true;
                dataGridView2.Columns["магнитныйВалНДС"].Visible = true;
            }
            if (checkBox7.Checked == false)//Лезвие очистки
            {
                dataGridView2.Columns["лезвиеОчисткиШтуки"].Visible = false;
                dataGridView2.Columns["лезвиеОчисткиЦена"].Visible = false;
                dataGridView2.Columns["лезвиеОчисткиСтоимость"].Visible = false;
                dataGridView2.Columns["лезвиеОчисткиНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["лезвиеОчисткиШтуки"].Visible = true;
                dataGridView2.Columns["лезвиеОчисткиЦена"].Visible = true;
                dataGridView2.Columns["лезвиеОчисткиСтоимость"].Visible = true;
                dataGridView2.Columns["лезвиеОчисткиНДС"].Visible = true;
            }
            if (checkBox8.Checked == false)//Чип
            {
                dataGridView2.Columns["чипШтуки"].Visible = false;
                dataGridView2.Columns["чипЦена"].Visible = false;
                dataGridView2.Columns["чипСтоимость"].Visible = false;
                dataGridView2.Columns["чипНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["чипШтуки"].Visible = true;
                dataGridView2.Columns["чипЦена"].Visible = true;
                dataGridView2.Columns["чипСтоимость"].Visible = true;
                dataGridView2.Columns["чипНДС"].Visible = true;
            }
            if (checkBox9.Checked == false)//Башинг
            {
                dataGridView2.Columns["башингШтуки"].Visible = false;
                dataGridView2.Columns["башингЦена"].Visible = false;
                dataGridView2.Columns["башингСтоимость"].Visible = false;
                dataGridView2.Columns["башингНДС"].Visible = false;
            }
            else
            {
                dataGridView2.Columns["башингШтуки"].Visible = true;
                dataGridView2.Columns["башингЦена"].Visible = true;
                dataGridView2.Columns["башингСтоимость"].Visible = true;
                dataGridView2.Columns["башингНДС"].Visible = true;
            }
        }
        public void StartChekerFormatTable(int rowind)
        {

            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "UPDATE Тендер SET [Картридж Штуки]=@Kartrid, [Тонер Штуки]=@Toner, [Барабан Штуки]=@Baraban, [Ракель Штуки]=@Rakel, [PCR Штуки]=@PCR, [Магнитный вал Штуки]=@MagVal, [Лезвие очистки Штуки]=@LezOch, [Чип Штуки]=@Chip, [Башинг Штуки]=@Bashing  WHERE Код = @KodTender";
            OleDbConnection conn1 = new OleDbConnection(ConnectionString);
            OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

            updataeCommand1.Parameters.AddWithValue("Kartrid", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);

            if (checkBox2.Checked)
                updataeCommand1.Parameters.AddWithValue("Toner", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("Toner", 0);
            if (checkBox4.Checked)
                updataeCommand1.Parameters.AddWithValue("Baraban", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("Baraban", 0);
            if (checkBox5.Checked)
                updataeCommand1.Parameters.AddWithValue("Rakel", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("Rakel", 0);
            if (checkBox1.Checked)
                updataeCommand1.Parameters.AddWithValue("PCR", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("PCR", 0);
            if (checkBox6.Checked)
                updataeCommand1.Parameters.AddWithValue("MagVal", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("MagVal", 0);
            if (checkBox7.Checked)
                updataeCommand1.Parameters.AddWithValue("LezOch", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("LezOch", 0);
            if (checkBox8.Checked)
                updataeCommand1.Parameters.AddWithValue("Chip", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("Chip", 0);
            if (checkBox9.Checked)
                updataeCommand1.Parameters.AddWithValue("Bashing", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
            else
                updataeCommand1.Parameters.AddWithValue("Bashing", 0);

            updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

            conn1.Open();
            int a = updataeCommand1.ExecuteNonQuery();
            conn1.Close();
        }

        public void PriceCheker()
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                dataGridView2.Rows[i].Cells["общаяНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["общаСумма"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["тонерНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["тонерСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["барабанНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["барабанСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["ракельНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["ракельСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["pCRНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["pCRСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["магнитныйВалНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["магнитныйВалСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["лезвиеОчисткиНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["лезвиеОчисткиСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["чипНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["чипСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
                dataGridView2.Rows[i].Cells["башингНДС"].Value = Convert.ToDouble(dataGridView2.Rows[i].Cells["башингСтоимость"].Value) * Convert.ToDouble(numericUpDown1.Value);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int FirstGrid2RowsCounter = dataGridView2.RowCount;
            таблица_для_ТендераTableAdapter.Update(заправка1DataSet);
            this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);

            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = " INSERT INTO Тендер ( [Название картриджа], [Тонер Масса], [Тонер Цена], [Барабан Цена], [Ракель Цена], [PCR Цена], [Магнитный вал Цена], [Лезвие очистки Цена], [Чип Цена], [Башинг Цена] ) "
                               + " SELECT Картридж.[Название картриджа], Картридж.[Масса тонера], Тонер.Цена, Барабан.Цена, Ракель.Цена, PCR.Цена, [Магнитный вал].Цена, [Лезвие очистки].Цена, Чип.Цена, Башинг.Цена "
                               + " FROM Башинг INNER JOIN(Чип INNER JOIN (Тонер INNER JOIN (Ракель INNER JOIN (Барабан INNER JOIN (PCR INNER JOIN ([Лезвие очистки] INNER JOIN ([Магнитный вал] INNER JOIN (Картридж INNER JOIN[Таблица для Тендера] ON Картридж.Код = [Таблица для Тендера].Индекс) ON[Магнитный вал].Код = Картридж.[Магнитный вал]) ON[Лезвие очистки].Код = Картридж.[Лезвие очистки]) ON PCR.Код = Картридж.PCR) ON Барабан.Код = Картридж.Барабан) ON Ракель.Код = Картридж.Ракель) ON Тонер.Код = Картридж.Тонер) ON Чип.Код = Картридж.Чип) ON Башинг.Код = Картридж.Башинг "
                               + " WHERE ((([Таблица для Тендера].Выбор)=True));";

            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand insertCommand = new OleDbCommand(commandText, conn);
            conn.Open();
            int PostGrid2RowsCounter = insertCommand.ExecuteNonQuery() + FirstGrid2RowsCounter;
            conn.Close();
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            TableCheker();
            for (; FirstGrid2RowsCounter < PostGrid2RowsCounter; FirstGrid2RowsCounter++)
            {
                StartChekerFormatTable(FirstGrid2RowsCounter);
            }
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            if (checkBox3.Checked)
            {
                string commandText2 = "UPDATE [Таблица для Тендера] SET [Таблица для Тендера].Выбор = False";
                insertCommand.CommandText = commandText2;
                conn.Open();
                insertCommand.ExecuteNonQuery();
                conn.Close();
                this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
            }
            PriceCheker();
        }  



        private void button2_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "DELETE * FROM Тендер;";
            OleDbConnection conn = new OleDbConnection(ConnectionString);
            OleDbCommand DeleteCommand = new OleDbCommand(commandText, conn);
            conn.Open();
            DeleteCommand.ExecuteNonQuery();
            conn.Close();
            this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int rowind = e.RowIndex;
            int columind = e.ColumnIndex;
            if ((rowind >= 0 && columind >= 0))
            {
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "";

                if (columind == 3)
                {

                    commandText = "UPDATE Тендер SET [Картридж Штуки]=@Kartrid, [Тонер Штуки]=@Toner, [Барабан Штуки]=@Baraban, [Ракель Штуки]=@Rakel, [PCR Штуки]=@PCR, [Магнитный вал Штуки]=@MagVal, [Лезвие очистки Штуки]=@LezOch, [Чип Штуки]=@Chip  WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                    updataeCommand1.Parameters.AddWithValue("Kartrid", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);

                    if (checkBox2.Checked)
                        updataeCommand1.Parameters.AddWithValue("Toner", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("Toner", 0);
                    if (checkBox4.Checked)
                        updataeCommand1.Parameters.AddWithValue("Baraban", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("Baraban", 0);
                    if (checkBox5.Checked)
                        updataeCommand1.Parameters.AddWithValue("Rakel", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("Rakel", 0);
                    if (checkBox1.Checked)
                        updataeCommand1.Parameters.AddWithValue("PCR", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("PCR", 0);
                    if (checkBox6.Checked)
                        updataeCommand1.Parameters.AddWithValue("MagVal", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("MagVal", 0);
                    if (checkBox7.Checked)
                        updataeCommand1.Parameters.AddWithValue("LezOch", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("LezOch", 0);
                    if (checkBox8.Checked)
                        updataeCommand1.Parameters.AddWithValue("Chip", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                    else
                        updataeCommand1.Parameters.AddWithValue("Chip", 0);

                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["код"].Value);

                    conn1.Open();
                    int a = updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 5)
                {
                    commandText = "UPDATE Тендер SET [Тонер Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 7)
                {
                    commandText = "UPDATE Тендер SET [Тонер Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 10)
                {
                    commandText = "UPDATE Тендер SET [Барабан Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 11)
                {
                    commandText = "UPDATE Тендер SET [Барабан Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 14)
                {
                    commandText = "UPDATE Тендер SET [Ракель Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 15)
                {
                    commandText = "UPDATE Тендер SET [Ракель Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 18)
                {
                    commandText = "UPDATE Тендер SET [PCR Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 19)
                {
                    commandText = "UPDATE Тендер SET [PCR Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 22)
                {
                    commandText = "UPDATE Тендер SET [Магнитный вал Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 23)
                {
                    commandText = "UPDATE Тендер SET [Магнитный вал Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 26)
                {
                    commandText = "UPDATE Тендер SET [Лезвие очистки Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 27)
                {
                    commandText = "UPDATE Тендер SET [Лезвие очистки Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 30)
                {
                    commandText = "UPDATE Тендер SET [Чип Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 31)
                {
                    commandText = "UPDATE Тендер SET [Чип Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 34)
                {
                    commandText = "UPDATE Тендер SET [Башинг Штуки]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
                if (columind == 35)
                {
                    commandText = "UPDATE Тендер SET [Башинг Цена]=@Varb WHERE Код = @KodTender";
                    OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                    OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
                    updataeCommand1.Parameters.AddWithValue("Varb", dataGridView2.Rows[rowind].Cells[columind].Value);
                    updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);
                    conn1.Open();
                    updataeCommand1.ExecuteNonQuery();
                    conn1.Close();
                    this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
                }
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Чип Штуки]=@Chip  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox8.Checked)
                    updataeCommand1.Parameters.AddWithValue("Chip", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("Chip", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Тонер Штуки]=@Toner  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox2.Checked)
                    updataeCommand1.Parameters.AddWithValue("Toner", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("Toner", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);

            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Барабан Штуки]=@Baraban  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox4.Checked)
                    updataeCommand1.Parameters.AddWithValue("Baraban", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("Baraban", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Ракель Штуки]=@Rakel  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox5.Checked)
                    updataeCommand1.Parameters.AddWithValue("Rakel", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("Rakel", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [PCR Штуки]=@PCR  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox1.Checked)
                    updataeCommand1.Parameters.AddWithValue("PCR", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("PCR", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Магнитный вал Штуки]=@MagVal  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox6.Checked)
                    updataeCommand1.Parameters.AddWithValue("MagVal", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("MagVal", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Лезвие очистки Штуки]=@LezOch  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox7.Checked)
                    updataeCommand1.Parameters.AddWithValue("LezOch", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("LezOch", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "UPDATE [Таблица для Тендера] SET Выбор=True";
            OleDbConnection conn1 = new OleDbConnection(ConnectionString);
            OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
            conn1.Open();
            updataeCommand1.ExecuteNonQuery();
            conn1.Close();
            this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
            string commandText = "UPDATE [Таблица для Тендера] SET Выбор=False";
            OleDbConnection conn1 = new OleDbConnection(ConnectionString);
            OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);
            conn1.Open();
            updataeCommand1.ExecuteNonQuery();
            conn1.Close();
            this.таблица_для_ТендераTableAdapter.Fill(this.заправка1DataSet.Таблица_для_Тендера);
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            TableCheker();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                int rowind = i;
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Заправка1.accdb";
                string commandText = "UPDATE Тендер SET [Башинг Штуки]=@Bashing  WHERE Код = @KodTender";
                OleDbConnection conn1 = new OleDbConnection(ConnectionString);
                OleDbCommand updataeCommand1 = new OleDbCommand(commandText, conn1);

                if (checkBox9.Checked)
                    updataeCommand1.Parameters.AddWithValue("Bashing", dataGridView2.Rows[rowind].Cells["картриджШтуки"].Value);
                else
                    updataeCommand1.Parameters.AddWithValue("Bashing", 0);

                updataeCommand1.Parameters.AddWithValue("KodTender", dataGridView2.Rows[rowind].Cells["Код"].Value);

                conn1.Open();
                int a = updataeCommand1.ExecuteNonQuery();
                conn1.Close();
                this.тендерTableAdapter.Fill(this.заправка1DataSet.Тендер);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView2.Columns["общаяНДС"].Visible = true;
            dataGridView2.Columns["общаСумма"].Visible = true;
            dataGridView2.Columns["тонерНДС"].Visible = true;
            dataGridView2.Columns["тонерСтоимость"].Visible = true;
            dataGridView2.Columns["барабанНДС"].Visible = true;
            dataGridView2.Columns["барабанСтоимость"].Visible = true;
            dataGridView2.Columns["ракельНДС"].Visible = true;
            dataGridView2.Columns["ракельСтоимость"].Visible = true;
            dataGridView2.Columns["pCRНДС"].Visible = true;
            dataGridView2.Columns["pCRСтоимость"].Visible = true;
            dataGridView2.Columns["магнитныйВалНДС"].Visible = true;
            dataGridView2.Columns["магнитныйВалСтоимость"].Visible = true;
            dataGridView2.Columns["лезвиеОчисткиНДС"].Visible = true;
            dataGridView2.Columns["лезвиеОчисткиСтоимость"].Visible = true;
            dataGridView2.Columns["чипНДС"].Visible = true;
            dataGridView2.Columns["чипСтоимость"].Visible = true;
            dataGridView2.Columns["башингНДС"].Visible = true;
            dataGridView2.Columns["башингСтоимость"].Visible = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView2.Columns["общаяНДС"].Visible = true;
            dataGridView2.Columns["общаСумма"].Visible = false;
            dataGridView2.Columns["тонерНДС"].Visible = true;
            dataGridView2.Columns["тонерСтоимость"].Visible = false;
            dataGridView2.Columns["барабанНДС"].Visible = true;
            dataGridView2.Columns["барабанСтоимость"].Visible = false;
            dataGridView2.Columns["ракельНДС"].Visible = true;
            dataGridView2.Columns["ракельСтоимость"].Visible = false;
            dataGridView2.Columns["pCRНДС"].Visible = true;
            dataGridView2.Columns["pCRСтоимость"].Visible = false;
            dataGridView2.Columns["магнитныйВалНДС"].Visible = true;
            dataGridView2.Columns["магнитныйВалСтоимость"].Visible = false;
            dataGridView2.Columns["лезвиеОчисткиНДС"].Visible = true;
            dataGridView2.Columns["лезвиеОчисткиСтоимость"].Visible = false;
            dataGridView2.Columns["чипНДС"].Visible = true;
            dataGridView2.Columns["чипСтоимость"].Visible = false;
            dataGridView2.Columns["башингНДС"].Visible = true;
            dataGridView2.Columns["башингСтоимость"].Visible = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView2.Columns["общаяНДС"].Visible = false;
            dataGridView2.Columns["общаСумма"].Visible = true;
            dataGridView2.Columns["тонерНДС"].Visible = false;
            dataGridView2.Columns["тонерСтоимость"].Visible = true;
            dataGridView2.Columns["барабанНДС"].Visible = false;
            dataGridView2.Columns["барабанСтоимость"].Visible = true;
            dataGridView2.Columns["ракельНДС"].Visible = false;
            dataGridView2.Columns["ракельСтоимость"].Visible = true;
            dataGridView2.Columns["pCRНДС"].Visible = false;
            dataGridView2.Columns["pCRСтоимость"].Visible = true;
            dataGridView2.Columns["магнитныйВалНДС"].Visible = false;
            dataGridView2.Columns["магнитныйВалСтоимость"].Visible = true;
            dataGridView2.Columns["лезвиеОчисткиНДС"].Visible = false;
            dataGridView2.Columns["лезвиеОчисткиСтоимость"].Visible = true;
            dataGridView2.Columns["чипНДС"].Visible = false;
            dataGridView2.Columns["чипСтоимость"].Visible = true;
            dataGridView2.Columns["башингНДС"].Visible = false;
            dataGridView2.Columns["башингСтоимость"].Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);
            ExcelApp.Columns.ColumnWidth = 16;
            ExcelApp.Cells[1, 1] = "Обща сумма";
            ExcelApp.Cells[1, 2] = "Обща НДС+";
            ExcelApp.Cells[1, 3] = "Картридж штуки";
            ExcelApp.Cells[1, 4] = "Название картриджа";
            ExcelApp.Cells[1, 5] = "Тонер Штуки";
            ExcelApp.Cells[1, 6] = "Тонер Масса";
            ExcelApp.Cells[1, 7] = "Тонер Цена";
            ExcelApp.Cells[1, 8] = "Тонер Стоимость";
            ExcelApp.Cells[1, 9] = "Тонер НДС+";
            ExcelApp.Cells[1, 10] = "Барабан Штуки";
            ExcelApp.Cells[1, 11] = "Барабан Цена";
            ExcelApp.Cells[1, 12] = "Барабан Стоимость";
            ExcelApp.Cells[1, 13] = "Барабан НДС +";
            ExcelApp.Cells[1, 14] = "Ракель Штуки";
            ExcelApp.Cells[1, 15] = "Ракель Цена";
            ExcelApp.Cells[1, 16] = "Ракель Стоимость";
            ExcelApp.Cells[1, 17] = "Ракель НДС+";
            ExcelApp.Cells[1, 18] = "PCR Штуки";
            ExcelApp.Cells[1, 19] = "PCR Цена";
            ExcelApp.Cells[1, 20] = "PCR Стоимость";
            ExcelApp.Cells[1, 21] = "PCR НДС+";
            ExcelApp.Cells[1, 22] = "Магнитный вал Штуки";
            ExcelApp.Cells[1, 23] = "Магнитный вал Цена";
            ExcelApp.Cells[1, 24] = "Магнитный вал Стоимость";
            ExcelApp.Cells[1, 25] = "Магнитный вал НДС+";
            ExcelApp.Cells[1, 26] = "Лезвие очистки Штуки";
            ExcelApp.Cells[1, 27] = "Лезвие очистки Цена";
            ExcelApp.Cells[1, 28] = "Лезвие очистки Стоимость";
            ExcelApp.Cells[1, 29] = "Лезвие очистки НДС+";
            ExcelApp.Cells[1, 30] = "Чип Штуки";
            ExcelApp.Cells[1, 31] = "Чип Цена";
            ExcelApp.Cells[1, 32] = "Чип Стоимость";
            ExcelApp.Cells[1, 33] = "Чип НДС+";
            ExcelApp.Cells[1, 34] = "Башинг Штуки";
            ExcelApp.Cells[1, 35] = "Башинг Цена";
            ExcelApp.Cells[1, 36] = "Башинг Стоимость";
            ExcelApp.Cells[1, 37] = "Башинг НДС+";

            for (int i = 1; i < dataGridView2.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    ExcelApp.Cells[j + 2, i] = (dataGridView2[i, j].Value).ToString();
                }
            }
            ExcelApp.Visible = true;
        }
    }
}
