using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ERC
{
    public partial class DeleteUKForm : Form
    {
        private static int id_uk_old;
        private static string uk_name_old;
        public DeleteUKForm(int id_uk, string uk_name)
        {
            InitializeComponent();
            id_uk_old = id_uk;
            uk_name_old = uk_name;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable t;
            if (checkBox1.Checked == true)
            {
                if (textBox2.Text == "") MessageBox.Show("В поле 'Название' ничего не указано!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    t = SQL.FillTable("select * from Управляющая_компания where Название = '" + textBox2.Text + "'");
                    if (t.Rows.Count != 0) MessageBox.Show("Управляющая компания с таким названием уже существует. Добавление невозможно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        t = SQL.FillTable("select MAX(ID_УК) from Управляющая_компания");
                        int max_value = Convert.ToInt32(t.Rows[0][0]) + 1;
                        SQL.ExecuteSQL("insert Управляющая_компания values ("+max_value+",'"+textBox2.Text+"')");
                        MessageBox.Show("Управляющая компания успешно добавлена. Её ID: " + max_value, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        SQL.ExecuteSQL("update Квартира set ID_УК = " + max_value + " where ID_УК = " + id_uk_old);
                        Program.form1.dataGridView3.DataSource = null;
                        SQL.ExecuteSQL("delete from Управляющая_компания where ID_УК = " + id_uk_old);
                        MessageBox.Show("Управляющая компания успешно удалена! Все квартиры переведены на баланс новой управляющей компании.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Program.form1.UpdateUKBindings();
                        Program.form1.dataGridView4.DataSource = null;
                        Close();
                    }
                }
            }
            else
            {
                int id_uk_new = 0;
                
                if(Int32.TryParse(textBox1.Text, out id_uk_new) == true)
                {
                    if (id_uk_new == id_uk_old) MessageBox.Show("Указанный ID совпадает с ID управляющей компании, которую вы собираетесь удалять!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        t = SQL.FillTable("select * from Управляющая_компания where ID_УК = " + textBox1.Text);
                        if (t.Rows.Count == 0) MessageBox.Show("Управляющей компании с таким ID не существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else
                        {
                            SQL.ExecuteSQL("update Квартира set ID_УК = " + t.Rows[0][0].ToString() + " where ID_УК = " + id_uk_old);
                            Program.form1.dataGridView3.DataSource = null;
                            SQL.ExecuteSQL("delete from Управляющая_компания where ID_УК = " + id_uk_old);
                            MessageBox.Show("Управляющая компания успешно удалена! Все квартиры переведены на баланс новой управляющей компании.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Program.form1.UpdateUKBindings();
                            Program.form1.dataGridView4.DataSource = null;
                            Close();
                        }
                    }
                }
                else
                {
                    string uk_name_new = textBox2.Text;
                    if (uk_name_new == uk_name_old) MessageBox.Show("Указанное вами название компании совпадает с названием управляющей компании, которую вы собираетесь удалять!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        t = SQL.FillTable("select * from Управляющая_компания where Название = '" + textBox2.Text + "'");
                        if (t.Rows.Count == 0) MessageBox.Show("Управляющей компании с таким названием не существует!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else
                        {
                            SQL.ExecuteSQL("update Квартира set ID_УК = " + t.Rows[0][0].ToString() + " where ID_УК = " + id_uk_old);
                            Program.form1.dataGridView3.DataSource = null;
                            SQL.ExecuteSQL("delete from Управляющая_компания where ID_УК = " + id_uk_old);
                            MessageBox.Show("Управляющая компания успешно удалена! Все квартиры переведены на баланс новой управляющей компании.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Program.form1.UpdateUKBindings();
                            Program.form1.dataGridView4.DataSource = null;
                            Close();
                        }
                    }
                }
            }
        }
    }
}
