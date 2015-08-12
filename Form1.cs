using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ATLExeCOMServerLib;
using System.Diagnostics;

namespace ERC
{
    public partial class Form1 : Form
    {
        ATLSimpleObjectSTA com_object = new ATLSimpleObjectSTA();
        private const string TemplatePath = "C:\\Templates\\";
        private const string SavedDocumentsPath = "C:\\Saved Documents\\";

        Word._Application oWord;
        public Form1()
        {
            DateTime dt = DateTime.Now;
            InitializeComponent();            
            comboBox1.Text = comboBox1.Items[0].ToString();
            comboBox2.Text = comboBox2.Items[0].ToString();
            UpdateUKBindings();
            comboBox4.Text = comboBox4.Items[0].ToString();
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        public void UpdateUKBindings()
        {
            comboBox3.Items.Clear();
            DataTable bindings = SQL.FillTable("select Название from Управляющая_компания order by Название");
            foreach (DataRow row in bindings.Rows)
            {
                comboBox3.Items.Add(row[0].ToString());
            }
            comboBox3.Text = comboBox3.Items[0].ToString();

        }

        private void AllInformingInterface()
        {
            DataTable debting_flats = SQL.FillTable("select distinct ID_квартиры from Долг order by ID_квартиры");
            foreach (DataRow row in debting_flats.Rows)
            {
                oWord = new Word.Application();
                DataTable flat_info = SQL.FillTable("select * from Квартира where ID_квартиры = " + row[0].ToString());
                DataTable tenant_info = SQL.FillTable("select * from Квартиросъёмщик where ID_квартиросъёмщика = " + flat_info.Rows[0][9].ToString());
                DataTable debt_info = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + row[0].ToString() + " order by Год, Месяц");

                Word._Document oDoc = oWord.Documents.Add(TemplatePath + "Задолженность_по_квартплате.dotx");
                oDoc.Bookmarks["ФИО"].Range.Text = tenant_info.Rows[0][3].ToString() + " " + tenant_info.Rows[0][4].ToString() + " "
                    + tenant_info.Rows[0][5].ToString();
                oDoc.Bookmarks["Город"].Range.Text = flat_info.Rows[0][1].ToString();
                oDoc.Bookmarks["Улица"].Range.Text = flat_info.Rows[0][2].ToString();
                oDoc.Bookmarks["Дом"].Range.Text = flat_info.Rows[0][3].ToString();
                oDoc.Bookmarks["Квартира"].Range.Text = flat_info.Rows[0][4].ToString();
                int i = 2;
                foreach (DataRow row_debt in debt_info.Rows)
                {
                    oDoc.Tables[1].Rows.Add();
                    oDoc.Tables[1].Cell(i, 1).Range.Text = row_debt[1].ToString();
                    oDoc.Tables[1].Cell(i, 2).Range.Text = row_debt[0].ToString();
                    oDoc.Tables[1].Cell(i, 3).Range.Text = row_debt[2].ToString();
                    i++;
                }
                DateTime dt = DateTime.Now;
                oDoc.SaveAs(SavedDocumentsPath + "ИОЗ-" + flat_info.Rows[0][1].ToString() + "-" + flat_info.Rows[0][2].ToString() + "-" +
                flat_info.Rows[0][3].ToString() + "-" + flat_info.Rows[0][4].ToString() + "-" + dt.Day + dt.Month + dt.Year + dt.Hour + dt.Minute +
                dt.Second + ".docx");
                oDoc.Close();
            }
        }

        private void RefreshDebtChart(int flat_id)
        {
            DataTable t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id);
            this.bindingSource1.DataSource = t;
            this.dataGridView1.DataSource = bindingSource1;
        }

        private void RefreshDebtChart(string city, string street, string house_number, string flat_number)
        {
            DataTable temp = SQL.FillTable("select ID_квартиры from Квартира where Город = '"+city+"' and Улица = '"+street
                +"' and Дом='"+house_number+"' and Номер_квартиры = '"+flat_number+"'");
            int flat_id = Convert.ToInt32(temp.Rows[0][0]);
            DataTable t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id);
            this.bindingSource1.DataSource = t;
            this.dataGridView1.DataSource = bindingSource1;
        }
        private int COMFormula(int sq_m, int livers, int coeff)
        {
            return (int)com_object.UtilitiesAccounting(coeff, livers, sq_m);
        }
        private void UtilitiesCount(int month, int year)
        {
            int Coefficient = 30;
            DataTable t = SQL.FillTable("select * from Квартира");
            foreach(DataRow row in t.Rows)
            {
                int to_pay = COMFormula(Convert.ToInt32(row[5]),Convert.ToInt32(row[6]),Coefficient);
                // формирование платёжки
                SQL.ExecuteSQL("insert Платёжка values ("+month+", "+year+", "+row[0]+", "+to_pay+")");
                // формирование оплаты
                int current_balance = Convert.ToInt32(row[7]);
                bool debt = false;
                if(current_balance - to_pay >= 0)
                {
                    SQL.ExecuteSQL("insert Оплата values (" + month + ", " + year + ", " + row[0] + ", " + to_pay + ")");
                    string cmd = "update Квартира set Баланс = "+ (current_balance - to_pay) +" where ID_квартиры = "
                        + row[0];
                    SQL.ExecuteSQL(cmd);
                }
                else
                {
                    debt = true;
                    SQL.ExecuteSQL("insert Оплата values (" + month + ", " + year + ", " + row[0] + ", " + current_balance + ")");
                    string cmd = "update Квартира set Баланс = 0 where ID_квартиры = " + row[0];
                    SQL.ExecuteSQL(cmd);
                }
                // формирование долга, если есть
                if(debt == true)
                {
                    string cmd = "insert Долг values (" + month + ", " + year + ", " + row[0] + ", " + (to_pay - current_balance) + ")";
                    SQL.ExecuteSQL(cmd);
                }
            }
            
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // пополнение баланса
            this.dataGridView1.DataSource = null;
            DataTable t;
            int flat_id;
            int balance = 0;
            int temp;
            bool isID = false; bool isDone = false;
            bool isNumber = Int32.TryParse(textBox4.Text, out temp);
            if (isNumber == true)
            {
                if (Int32.TryParse(textBox9.Text, out flat_id) == true)
                {
                    isID = true;
                    t = SQL.FillTable("select Баланс from Квартира where ID_квартиры = " + flat_id);
                    if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким ID не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        balance = Convert.ToInt32(t.Rows[0][0]) + Convert.ToInt32(textBox4.Text);
                        string cmd = "update Квартира set Баланс = "
                            + balance + "where ID_квартиры = " + flat_id;
                        SQL.ExecuteSQL(cmd);
                        isDone = true;
                        MessageBox.Show("Баланс успешно пополнен! Текущий баланс: " + balance + " рублей", "Информация");
                        this.dataGridView4.DataSource = null;
                    }

                }
                else
                {
                    t = SQL.FillTable("select Баланс, ID_квартиры from Квартира where Город = '" + comboBox1.Text + "' and Улица = '"
                        + textBox1.Text + "' and Дом = '" + textBox2.Text + "' and Номер_квартиры = '" + textBox3.Text + "'");
                    if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким адресом не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        balance = Convert.ToInt32(t.Rows[0][0]) + Convert.ToInt32(textBox4.Text);
                        string cmd = "update Квартира set Баланс = "
                            + balance + "where ID_квартиры = " + t.Rows[0][1];
                        SQL.ExecuteSQL(cmd);
                        isDone = true;
                        MessageBox.Show("Баланс успешно пополнен! Текущий баланс: " + balance + " рублей", "Информация");
                        this.dataGridView4.DataSource = null;
                    }
                }
                // списание долгов, если есть

                if (isDone == true)
                {
                    bool has_debts = false; bool fully_paid = true;
                    if (isID == false) flat_id = Convert.ToInt32(t.Rows[0][1]);
                    t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id
                        + "order by Год, Месяц");
                    if (t.Rows.Count != 0) has_debts = true;
                    foreach (DataRow row in t.Rows)
                    {
                        if (balance > 0)
                        {
                            if (balance - Convert.ToInt32(row[2]) >= 0)
                            {
                                SQL.ExecuteSQL("delete from Долг where Месяц = " + row[0] + " and Год = " + row[1]
                                    + " and ID_квартиры = " + flat_id);
                                string cmd = "update Оплата set Сумма = Сумма + " + row[2] + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id;
                                SQL.ExecuteSQL(cmd);
                                balance -= Convert.ToInt32(row[2]);
                            }
                            else
                            {
                                SQL.ExecuteSQL("update Долг set Сумма = Сумма - " + balance + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id);
                                SQL.ExecuteSQL("update Оплата set Сумма = Сумма + " + balance + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id);
                                balance = 0;
                                fully_paid = false;
                            }
                        }
                        else fully_paid = false;

                    }
                    if (has_debts == true)
                    {
                        SQL.ExecuteSQL("update Квартира set Баланс = " + balance + " where ID_квартиры = " + flat_id);
                        if (fully_paid == true)
                        {
                            MessageBox.Show("На счету данной квартиры имелись долги. Все они полностью списаны! Текущий баланс: " + balance + " рублей", "Информация");
                        }
                        else
                        {
                            t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id
                            + "order by Год, Месяц");
                            int sum = 0;
                            foreach (DataRow row in t.Rows)
                            {
                                sum += Convert.ToInt32(row[2]);
                            }
                            MessageBox.Show("На счету данной квартиры имеются долги. Они списаны частично!\nДля полного списания долгов необходимо пополнить баланс ещё на " + sum + " рублей! На данный момент на вашем счету остаётся 0 рублей", "Информация");
                        }
                    }
                }
            }
            else MessageBox.Show("Проверьте правильность ввода данных в поле 'Сумма к оплате'!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private Word._Document GetDoc(string path)
        {
            Word._Document oDoc = oWord.Documents.Add(path);
            SetTemplate(oDoc);
            return oDoc;
        }
        // Замена закладки SECONDNAME на данные введенные в textBox
        private void SetTemplate(Word._Document oDoc)
        {
            oDoc.Bookmarks["STREET"].Range.Text = textBox1.Text;
            oDoc.Bookmarks["HOUSENUMBER"].Range.Text = textBox2.Text;

            oDoc.Tables[1].Rows.Add();
            oDoc.Tables[1].Cell(2, 1).Range.Text = "Чистяков";
            oDoc.Tables[1].Cell(2, 2).Range.Text = "Павел";
            oDoc.Tables[1].Cell(2, 3).Range.Text = "Александрович";
            // если нужно заменять другие закладки, тогда копируем верхнюю строку изменяя на нужные параметры 

        }

        private void расчётКToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Now;
            string cmd = "select * from Расчёт_проведен where Месяц = " + date.Month
                 + " and Год = " + date.Year;
            DataTable t = new DataTable();
            t = SQL.FillTable(cmd);
            if (t.Rows.Count != 0)
            {
                MessageBox.Show("Расчёт за данный месяц уже проведён", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DialogResult dr = MessageBox.Show("Вы действительно хотите провести расчёт за этот месяц?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    
                    UtilitiesCount(date.Month, date.Year);

                    SQL.ExecuteSQL("INSERT Расчёт_проведен values (" + date.Month + "," + date.Year + ",1)");

                    MessageBox.Show("Расчёт проведён успешно!", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable t;
            int flat_id;
            if(Int32.TryParse(textBox9.Text, out flat_id) == true)
            {
                t = SQL.FillTable("select Баланс from Квартира where ID_квартиры = " + flat_id);
                if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким ID не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    MessageBox.Show("Баланс квартиры с ID " + flat_id + ": " + t.Rows[0][0] + " рублей", "Информация");
                }
            }
            else
            {
                t = SQL.FillTable("select Баланс from Квартира where Город = '" + comboBox1.Text + "' and Улица = '"
                    + textBox1.Text + "' and Дом = '"+ textBox2.Text +"' and Номер_квартиры = '"+textBox3.Text+"'");
                if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким адресом не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    MessageBox.Show("Баланс квартиры с таким адресом: " + t.Rows[0][0] + " рублей", "Информация");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable t;
            int flat_id;
            this.dataGridView1.DataSource = null;
            if(Int32.TryParse(textBox10.Text, out flat_id) == true)
            {
                t = SQL.FillTable("select * from Квартира where ID_квартиры = "+flat_id);
                if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким ID не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id);
                    if (t.Rows.Count == 0) MessageBox.Show("На счету данной квартиры долгов нет! Ничего не выведено!", "Информация");
                    else
                    {
                        RefreshDebtChart(flat_id);
                    }
                }
            }
            else
            {
                t = SQL.FillTable("select ID_квартиры from Квартира where Город = '" + comboBox4.Text + "' and Улица = '"+
                    textBox7.Text + "' and Дом = '"+textBox6.Text+"' and Номер_квартиры = '"+textBox5.Text+"'");
                if (t.Rows.Count == 0) { MessageBox.Show("Квартира с таким адресом не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                else
                {
                    flat_id = Convert.ToInt32(t.Rows[0][0]);
                    t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id);
                    if (t.Rows.Count == 0) MessageBox.Show("На счету данной квартиры долгов нет! Ничего не выведено!", "Информация");
                    else
                    {
                        RefreshDebtChart(comboBox4.Text, textBox7.Text, textBox6.Text, textBox5.Text);
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // пополнение баланса
            
            DataTable t;
            int flat_id;
            int balance = 0;
            int temp;
            bool isID = false; bool isDone = false;
            bool isNumber = Int32.TryParse(textBox8.Text, out temp);
            if (isNumber == true)
            {
                if (Int32.TryParse(textBox10.Text, out flat_id) == true)
                {
                    isID = true;
                    t = SQL.FillTable("select Баланс from Квартира where ID_квартиры = " + flat_id);
                    if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким ID не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        balance = Convert.ToInt32(t.Rows[0][0]) + Convert.ToInt32(textBox8.Text);
                        string cmd = "update Квартира set Баланс = "
                            + balance + "where ID_квартиры = " + flat_id;
                        SQL.ExecuteSQL(cmd);
                        isDone = true;
                        
                        MessageBox.Show("Баланс успешно пополнен! Текущий баланс: " + balance + " рублей", "Информация");
                        this.dataGridView4.DataSource = null;
                    }

                }
                else
                {
                    t = SQL.FillTable("select Баланс, ID_квартиры from Квартира where Город = '" + comboBox4.Text + "' and Улица = '"
                        + textBox7.Text + "' and Дом = '" + textBox6.Text + "' and Номер_квартиры = '" + textBox5.Text + "'");
                    if (t.Rows.Count == 0) MessageBox.Show("Квартира с таким адресом не найдена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        balance = Convert.ToInt32(t.Rows[0][0]) + Convert.ToInt32(textBox8.Text);
                        string cmd = "update Квартира set Баланс = "
                            + balance + "where ID_квартиры = " + t.Rows[0][1];
                        SQL.ExecuteSQL(cmd);
                        isDone = true;
                        
                        MessageBox.Show("Баланс успешно пополнен! Текущий баланс: " + balance + " рублей", "Информация");
                        this.dataGridView4.DataSource = null;
                        
                    }
                }
                // списание долгов, если есть

                if (isDone == true)
                {
                    bool has_debts = false; bool fully_paid = true;
                    if (isID == false) flat_id = Convert.ToInt32(t.Rows[0][1]);
                    t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id
                        + "order by Год, Месяц");
                    if (t.Rows.Count != 0) has_debts = true;
                    foreach (DataRow row in t.Rows)
                    {
                        if (balance > 0)
                        {
                            if (balance - Convert.ToInt32(row[2]) >= 0)
                            {
                                SQL.ExecuteSQL("delete from Долг where Месяц = " + row[0] + " and Год = " + row[1]
                                    + " and ID_квартиры = " + flat_id);
                                string cmd = "update Оплата set Сумма = Сумма + " + row[2] + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id;
                                SQL.ExecuteSQL(cmd);
                                balance -= Convert.ToInt32(row[2]);
                            }
                            else
                            {
                                SQL.ExecuteSQL("update Долг set Сумма = Сумма - " + balance + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id);
                                SQL.ExecuteSQL("update Оплата set Сумма = Сумма + " + balance + " where " +
                                    "Месяц = " + row[0] + " and Год = " + row[1] + " and ID_квартиры = " + flat_id);
                                balance = 0;
                                fully_paid = false;
                            }
                        }
                        else fully_paid = false;

                    }
                    if (has_debts == true)
                    {
                        
                        SQL.ExecuteSQL("update Квартира set Баланс = " + balance + " where ID_квартиры = " + flat_id);
                        if (fully_paid == true)
                        {
                            this.dataGridView1.DataSource = null;
                            MessageBox.Show("На счету данной квартиры имелись долги. Все они полностью списаны! Текущий баланс: " + balance + " рублей", "Информация");
                        }
                        else
                        {
                            t = SQL.FillTable("select Месяц, Год, Сумма from Долг where ID_квартиры = " + flat_id
                            + "order by Год, Месяц");
                            int sum = 0;
                            foreach (DataRow row in t.Rows)
                            {
                                sum += Convert.ToInt32(row[2]);
                            }
                            RefreshDebtChart(flat_id);
                            MessageBox.Show("На счету данной квартиры имеются долги. Они списаны частично!\nДля полного списания долгов необходимо пополнить баланс ещё на " + sum + " рублей! На данный момент на вашем счету остаётся 0 рублей", "Информация");
                        }
                    }
                    else this.dataGridView1.DataSource = null;
                    
                }
            }
            else MessageBox.Show("Проверьте правильность ввода данных в поле 'Сумма к оплате'!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox11.Text != "" && textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "" &&
                textBox15.Text != "")
            {
                DataTable if_exist = SQL.FillTable("select * from Квартиросъёмщик where Серия_паспорта='"+textBox14.Text
                    +"' and Номер_паспорта='"+textBox15.Text+"'");
                if (if_exist.Rows.Count != 0) MessageBox.Show("Квартиросъёмщик с такой серией и номером паспорта уже существует! Добавление невозможно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    string message = "Вы добавляете квартиросъёмщика. Проверьте правильность ввода данных: \n" +
                        "Фамилия: " + textBox11.Text + "\n" +
                        "Имя: " + textBox12.Text + "\n" +
                        "Отчество: " + textBox13.Text + "\n" +
                        "Серия паспорта: " + textBox14.Text + "\n" +
                        "Номер паспорта: " + textBox15.Text + "\n" +
                        "Если всё введено верно, нажмите кнопку 'OK' для добавления квартиросъёмщика";
                    DialogResult dr = MessageBox.Show(message, "Информация", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.OK)
                    {
                        int max_value = 0;
                        DataTable if_null = SQL.FillTable("select * from Квартиросъёмщик");
                        if (if_null.Rows.Count == 0) max_value = 1;
                        else
                        {
                            DataTable max = SQL.FillTable("select MAX(ID_квартиросъёмщика) from Квартиросъёмщик");
                            max_value = Convert.ToInt32(max.Rows[0][0]) + 1;
                        }
                        SQL.ExecuteSQL("insert Квартиросъёмщик values ("+max_value+",'" + textBox14.Text + "','" +
                            textBox15.Text + "','" + textBox11.Text + "','" + textBox12.Text + "','" + textBox13.Text + "')");
                        MessageBox.Show("Квартиросъёмщик успешно добавлен! Его идентификационный номер: "+max_value+"", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Некоторые из полей остались пустыми! Операция невозможна!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataTable t;
            this.dataGridView2.DataSource = null;
            int tenant_id;
            if (Int32.TryParse(textBox22.Text, out tenant_id) == true)
            {
                t = SQL.FillTable("select * from Квартиросъёмщик where ID_квартиросъёмщика = " + tenant_id);
                if (t.Rows.Count == 0) MessageBox.Show("Квартиросъёмщика с таким ID не существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource2.DataSource = t;
                    this.dataGridView2.DataSource = bindingSource2;
                    this.dataGridView2.ClearSelection();
                }
            }
            else
            {
                t = SQL.FillTable("select * from Квартиросъёмщик where Серия_паспорта = '" + textBox14.Text + "'"
                    + " and Номер_паспорта = '" + textBox15.Text + "'");
                if (t.Rows.Count == 0) MessageBox.Show("Квартиросъёмщика с такими данными не существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource2.DataSource = t;
                    this.dataGridView2.DataSource = bindingSource2;
                    this.dataGridView2.ClearSelection();
                }
            }
        }

        private void dgv2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView2.Rows[rowIndex];
            textBox22.Text = row.Cells[0].Value.ToString();
            textBox14.Text = row.Cells[1].Value.ToString();
            textBox15.Text = row.Cells[2].Value.ToString();
            textBox11.Text = row.Cells[3].Value.ToString();
            textBox12.Text = row.Cells[4].Value.ToString();
            textBox13.Text = row.Cells[5].Value.ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView2.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView2.Rows[index];
                if(textBox11.Text != "" && textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "" &&
                textBox15.Text != "")
                {                    
                        string message = "Вы изменяете данные квартиросъёмщика с ID: "+row.Cells[0].Value.ToString()+". Старые данные: \n\n" +
                       "Фамилия: " + row.Cells[3].Value.ToString() + "\n" +
                       "Имя: " + row.Cells[4].Value.ToString() + "\n" +
                       "Отчество: " + row.Cells[5].Value.ToString() + "\n" +
                       "Серия паспорта: " + row.Cells[1].Value.ToString() +"\n" +
                       "Номер паспорта: " + row.Cells[2].Value.ToString() + "\n\n" +
                       "Проверьте правильность ввода новых данных: \n\n" +
                       "Фамилия: " + textBox11.Text + "\n" +
                       "Имя: " + textBox12.Text + "\n" +
                       "Отчество: " + textBox13.Text + "\n" +
                       "Серия паспорта: " + textBox14.Text + "\n" +
                       "Номер паспорта: " + textBox15.Text + "\n\n" +
                       "Если всё введено верно, нажмите кнопку 'OK' для изменения данных квартиросъёмщика";
                        DialogResult dr = MessageBox.Show(message, "Информация", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                        if (dr == DialogResult.OK)
                        {
                            SQL.ExecuteSQL("update Квартиросъёмщик set Серия_паспорта = '"+ textBox14.Text +"', Номер_паспорта = '"+textBox15.Text+
                                "', Фамилия = '"+textBox11.Text+"', Имя = '"+textBox12.Text+"'," +
                                " Отчество = '"+textBox13.Text+"' where ID_Квартиросъёмщика = " + row.Cells[0].Value.ToString());
                            MessageBox.Show("Данные квартиросъёмщика с ID "+row.Cells[0].Value.ToString()+" успешно изменены", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            DataTable refresh = SQL.FillTable("select * from Квартиросъёмщик where ID_квартиросъёмщика = " + row.Cells[0].Value.ToString());
                            this.bindingSource2.DataSource = refresh;
                            this.dataGridView2.DataSource = bindingSource2;
                            this.dataGridView2.ClearSelection();
                            this.dataGridView4.DataSource = null;
                        }
                }
                else
                {
                    MessageBox.Show("Некоторые из полей остались пустыми! Операция невозможна!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView2.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView2.Rows[index];
                DataTable tenant_flats = SQL.FillTable("select * from Квартира where ID_Квартиросъёмщика = " + row.Cells[0].Value.ToString());
                if (tenant_flats.Rows.Count != 0) MessageBox.Show("На балансе этого квартиросъёмщика есть квартиры. Невозможно удалить этого квартиросъёмщика, пока его квартиры не будут переоформлены на другое лицо", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    string message = "ВНИМАНИЕ! Вы удаляете квартиросъёмщика с ID: " + row.Cells[0].Value.ToString() + ". Его данные: \n" +
                       "Фамилия: " + row.Cells[3].Value.ToString() + "\n" +
                       "Имя: " + row.Cells[4].Value.ToString() + "\n" +
                       "Отчество: " + row.Cells[5].Value.ToString() + "\n" +
                       "Серия паспорта: " + row.Cells[1].Value.ToString() + "\n" +
                       "Номер паспорта: " + row.Cells[2].Value.ToString() + "\n\n" +                       
                       "Нажмите кнопку 'OK' для удаления квартиросъёмщика";
                    DialogResult dr = MessageBox.Show(message, "Информация", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if(dr == DialogResult.OK)
                    {
                        SQL.ExecuteSQL("delete from Квартиросъёмщик where ID_квартиросъёмщика = " + row.Cells[0].Value.ToString());
                        MessageBox.Show("Квартиросъёмщик с ID " + row.Cells[0].Value.ToString() + " успешно удалён", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.dataGridView2.DataSource = null;
                        textBox11.Text = ""; textBox12.Text = ""; textBox13.Text = ""; textBox14.Text = ""; textBox15.Text = ""; textBox22.Text = "";
                        this.dataGridView4.DataSource = null;

                    }
                }
                
            }
            catch(ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox16.Text == "") MessageBox.Show("Название управляющей компании не введено", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataTable if_exist = SQL.FillTable("select * from Управляющая_компания where Название = '" + textBox16.Text + "'");
                if (if_exist.Rows.Count != 0) MessageBox.Show("Управляющая компания с таким названием уже существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    int max_value = 0;
                    DataTable if_null = SQL.FillTable("select * from Управляющая_компания");
                    if (if_null.Rows.Count == 0) max_value = 1;
                    else
                    {
                        DataTable max = SQL.FillTable("select MAX(ID_УК) from Управляющая_компания");
                        max_value = Convert.ToInt32(max.Rows[0][0]) + 1;
                    }
                    SQL.ExecuteSQL("insert Управляющая_компания values ("+max_value+", '"+textBox16.Text+"')");
                    MessageBox.Show("Управляющая компания успешно добавлена! Её идентификационный номер: " + max_value + "", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateUKBindings();
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DataTable t;
            this.dataGridView3.DataSource = null;
            int uk_id;
            if (Int32.TryParse(textBox23.Text, out uk_id))
            {
                t = SQL.FillTable("select * from Управляющая_компания where ID_УК = " + uk_id);
                if (t.Rows.Count == 0) MessageBox.Show("Управляющей компании с таким ID не найдено!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource3.DataSource = t;
                    this.dataGridView3.DataSource = bindingSource3;
                    this.dataGridView3.ClearSelection();
                }
            }
            else
            {
                t = SQL.FillTable("select * from Управляющая_компания where Название = '"+textBox16.Text+"'");
                if (t.Rows.Count == 0) MessageBox.Show("Управляющая компания с таким названием не найдена", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource3.DataSource = t;
                    this.dataGridView3.DataSource = bindingSource3;
                    this.dataGridView3.ClearSelection();
                }
            }
        }

        private void dgv3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView3.Rows[rowIndex];
            textBox16.Text = row.Cells[1].Value.ToString();
            textBox23.Text = row.Cells[0].Value.ToString();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView3.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView3.Rows[index];
                DataTable t = SQL.FillTable("select * from Управляющая_компания where Название = '" + textBox16.Text + "'");
                if (t.Rows.Count != 0) MessageBox.Show("Управляющая компания с таким названием уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    SQL.ExecuteSQL("update Управляющая_компания set Название = '"+textBox16.Text+"' where ID_УК = " + textBox23.Text);
                    t = SQL.FillTable("select * from Управляющая_компания where ID_УК = " + textBox23.Text);
                    this.bindingSource3.DataSource = t;
                    this.dataGridView3.DataSource = bindingSource3;
                    this.dataGridView3.ClearSelection();
                    MessageBox.Show("Название управляющей компании успешно изменено!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateUKBindings();
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView3.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView3.Rows[index];
                DataTable t = SQL.FillTable("select * from Квартира where ID_УК = " + row.Cells[0].Value.ToString());
                if (t.Rows.Count != 0)
                {
                    DeleteUKForm new_form = new DeleteUKForm(Convert.ToInt32(row.Cells[0].Value), row.Cells[1].Value.ToString());
                    new_form.ShowDialog();
                }
                else
                {
                    SQL.ExecuteSQL("delete from Управляющая_компания where ID_УК = " + row.Cells[0].Value.ToString());
                    this.dataGridView3.DataSource = null;
                    MessageBox.Show("Управляющая компания успешно удалена!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateUKBindings();
                    this.dataGridView4.DataSource = null;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            DataTable t;
            int flat_id;
            if(Int32.TryParse(textBox17.Text, out flat_id))
            {
                t = SQL.FillTable("select * from Квартира where ID_квартиры = " + flat_id);
                if (t.Rows.Count == 0) MessageBox.Show("Квартиры с таким ID не найдено!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource4.DataSource = t;
                    this.dataGridView4.DataSource = bindingSource4;
                    this.dataGridView4.ClearSelection();
                }
            }
            else
            {
                t = SQL.FillTable("select * from Квартира where Город = '" + comboBox2.Text + "' and Улица = '"+textBox20.Text+"' and Дом = '"+
                    textBox19.Text+"' and Номер_квартиры = '"+textBox18.Text+"'");
                if (t.Rows.Count == 0) MessageBox.Show("Квартиры с таким адресом не найдено!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    this.bindingSource4.DataSource = t;
                    this.dataGridView4.DataSource = bindingSource4;
                    this.dataGridView4.ClearSelection();
                }
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (textBox24.Text == "") { MessageBox.Show("Квартиросъёмщик не выбран!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            else
            {
                DataTable t = SQL.FillTable("select * from Квартиросъёмщик where ID_квартиросъёмщика = " + textBox24.Text);
                string message = "Информация о квартиросъёмщике с ID " + textBox24.Text + "\n\n" +
                    "Фамилия: " + t.Rows[0][3].ToString() + "\n" +
                    "Имя: " + t.Rows[0][4].ToString() + "\n" +
                    "Отчество: " + t.Rows[0][5].ToString() + "\n" +
                    "Серия паспорта: " + t.Rows[0][1].ToString() + "\n" +
                    "Номер паспорта: " + t.Rows[0][2].ToString();
                MessageBox.Show(message, "Информация о квартиросъёмщике", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            TenantChoiseForm tcf = new TenantChoiseForm();
            tcf.ShowDialog();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("Вы действительно хотите выйти?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.No) e.Cancel = true;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataTable t;
            if (comboBox2.Text != "" && textBox20.Text != "" && textBox19.Text != "" && textBox18.Text != "" && textBox21.Text != "" &&
                textBox25.Text != "" && comboBox3.Text != "" && textBox24.Text != "")
            {
                t = SQL.FillTable("select * from Квартира where Город = '"+comboBox2.Text+"' and Улица = '"+textBox20.Text+"' and Дом = '"+
                    textBox19.Text+"' and Номер_квартиры = '"+textBox18.Text+"'");
                if (t.Rows.Count != 0) MessageBox.Show("Квартира с такими данными уже существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    //string message = "Вы добавляете квартиру. Проверьте правильность ввода данных:\n\n" +
                    //"Город: " + comboBox3.Text + "\n" +
                    //"Улица: " + 
                    /*DialogResult dr = MessageBox.Show("Сделай диалог с данными!!!", "Информация", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    if (dr == DialogResult.Yes)
                    {*/
                        int max_value = 0;
                        t = SQL.FillTable("select * from Квартира");
                        if (t.Rows.Count == 0) max_value = 1;
                        else
                        {
                            t = SQL.FillTable("select MAX(ID_квартиры) from Квартира");
                            max_value = Convert.ToInt32(t.Rows[0][0]) + 1;
                        }
                        t = SQL.FillTable("select ID_УК from Управляющая_компания where Название = '" + comboBox3.Text + "'");
                        int uk_id = Convert.ToInt32(t.Rows[0][0]);
                        SQL.ExecuteSQL("insert Квартира values (" + max_value + ",'" + comboBox2.Text + "','" + textBox20.Text + "','" + textBox19.Text + "','" +
                            textBox18.Text + "'," + textBox21.Text + "," + textBox25.Text + ",0," + uk_id + "," + textBox24.Text + ")");
                        MessageBox.Show("Квартира успешно добавлена! Её идентификационный номер: " + max_value, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //}                    
                }
            }
            else
            {
                MessageBox.Show("Некоторые из полей не введены! Продолжение операции невозможно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgv4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView4.Rows[rowIndex];
            textBox17.Text = row.Cells[0].Value.ToString();
            comboBox2.Text = row.Cells[1].Value.ToString();
            textBox20.Text = row.Cells[2].Value.ToString();
            textBox19.Text = row.Cells[3].Value.ToString();
            textBox18.Text = row.Cells[4].Value.ToString();
            textBox21.Text = row.Cells[5].Value.ToString();
            textBox25.Text = row.Cells[6].Value.ToString();
            DataTable t = SQL.FillTable("select Название from Управляющая_компания where ID_УК = " + row.Cells[8].Value.ToString());
            comboBox3.Text = t.Rows[0][0].ToString();
            textBox24.Text = row.Cells[9].Value.ToString();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView4.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView4.Rows[index];
                DataTable if_exist = SQL.FillTable("select * from Квартира where Город = '"+comboBox2.Text+"' and Улица = '"+textBox20.Text+
                    "' and Дом = '"+textBox19.Text+"' and Номер_квартиры = '"+textBox18.Text+"'");
                if (if_exist.Rows.Count != 0) MessageBox.Show("Квартира с таким адресом уже существует! Изменение данных невозможно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    DataTable uk = SQL.FillTable("select ID_УК from Управляющая_компания where Название = '" + comboBox3.Text + "'");
                    int uk_id = Convert.ToInt32(uk.Rows[0][0]);
                    SQL.ExecuteSQL("update Квартира set Город = '"+comboBox2.Text+"', Улица = '"+textBox20.Text+"', Дом = '"+textBox19.Text
                        +"', Номер_квартиры = '"+textBox18.Text+"', Площадь_квартиры = "+textBox21.Text+", "+
                        "Количество_проживающих = "+textBox25.Text+", ID_УК = "+uk_id+", ID_квартиросъёмщика = "+textBox24.Text+" where ID_квартиры = " + textBox17.Text);
                    MessageBox.Show("Изменение квартиры с ID " + textBox17.Text + " успешно завершено!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.dataGridView4.DataSource = null;
                }

            }
            catch(ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView4.SelectedRows[0].Index;
                DataGridViewRow row = dataGridView4.Rows[index];
                DataTable t = SQL.FillTable("select * from Долг where ID_квартиры = " + row.Cells[0].Value.ToString());
                if (t.Rows.Count != 0) MessageBox.Show("Невозможно удалить квартиру, пока на её балансе есть долги!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    /*DialogResult dr = MessageBox.Show("Сделай диалог с данными!", "Информация", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (dr == DialogResult.OK)
                    {*/
                        //!!!
                        SQL.ExecuteSQL("delete from Оплата where ID_квартиры = " + row.Cells[0].Value.ToString());
                        SQL.ExecuteSQL("delete from Платёжка where ID_квартиры = " + row.Cells[0].Value.ToString());
                        //!!!
                        SQL.ExecuteSQL("delete from Квартира where ID_квартиры = " + row.Cells[0].Value.ToString());
                        
                        MessageBox.Show("Квартира успешно удалена! К выдаче квартиросъёмщику " + row.Cells[7].Value.ToString() + " рублей", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.dataGridView4.DataSource = null;
                    //}
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("Никакая строка не выделена!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void информированиеОЗадолженностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable t = SQL.FillTable("select * from Долг");
            if (t.Rows.Count != 0)
            {
                AllInformingInterface();
                MessageBox.Show("Все документы сформированы!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else 
            {
                MessageBox.Show("Долгов нет. Никакие документы не сформированы.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void выходToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private void отчётПоКвартплатамЗаМесяцToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MonthReportForm form_month = new MonthReportForm();
            form_month.ShowDialog();
        }

        private void отчётПоКвартплатамЗаГодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            YearReportForm form_year = new YearReportForm();
            form_year.ShowDialog();
        }

        private void месячныеОтчётаПоУправляющимКомпаниямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UKReportingForm ukform = new UKReportingForm();
            ukform.ShowDialog();
        }

        private void рекламаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(@"C:\MenshikovaLab\Commerical\Реклама.pptx",
            Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue);
        }

        private void инструкцияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\MenshikovaLab\Tutorial\Руководство_пользователя.docx");
        }

        
    }
    // далее мусор
    // SQL.ExecuteSQL("INSERT Квартиросъёмщик values (1, '113', '11111', 'Чистяков', 'Павел', 'Александрович')");
    // Word._Document oDoc = GetDoc(Environment.CurrentDirectory + "\\Templates\\SimpleTemplate2.dotx");
    // oDoc.SaveAs(FileName: Environment.CurrentDirectory + "\\Saved Documents\\333.docx");
    // oDoc.Close();

}