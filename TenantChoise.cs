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
    public partial class TenantChoiseForm : Form
    {
        public TenantChoiseForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable t;
            int tenant_id = 0;
            if (Int32.TryParse(textBox1.Text, out tenant_id) == true)
            {
                t = SQL.FillTable("select * from Квартиросъёмщик where ID_квартиросъёмщика = " + tenant_id);
                if (t.Rows.Count == 0) MessageBox.Show("Квартиросъёмщик с таким ID не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else 
                {
                    Program.form1.textBox24.Text = tenant_id.ToString();
                    string message = "Вы выбрали квартиросъёмщика с ID " + tenant_id + ". Его данные:\n\n" +
                        "Фамилия: " +t.Rows[0][3].ToString()+ "\n" +
                        "Имя: " + t.Rows[0][4].ToString() + "\n"+
                        "Отчество: " + t.Rows[0][5].ToString() + "\n" +
                        "Серия паспорта: " + t.Rows[0][1].ToString() + "\n" +
                        "Номер паспорта: " + t.Rows[0][2].ToString();
                    MessageBox.Show(message, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();
                }
            }
            else
            {
                t = SQL.FillTable("select * from Квартиросъёмщик where Фамилия = '"+textBox2.Text+"' and Имя = '"+textBox3.Text+"' and Отчество = '"
                    + textBox4.Text+"' and Серия_паспорта = '"+textBox5.Text+"' and Номер_паспорта='"+textBox6.Text+"'");
                if (t.Rows.Count == 0) MessageBox.Show("Квартиросъёмщик с такими данными не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    tenant_id = Convert.ToInt32(t.Rows[0][0]);
                    Program.form1.textBox24.Text = tenant_id.ToString();
                    string message = "Вы выбрали квартиросъёмщика с ID " + tenant_id + ". Его данные:\n\n" +
                        "Фамилия: " + t.Rows[0][3].ToString() + "\n" +
                        "Имя: " + t.Rows[0][4].ToString() + "\n" +
                        "Отчество: " + t.Rows[0][5].ToString() + "\n" +
                        "Серия паспорта: " + t.Rows[0][1].ToString() + "\n" +
                        "Номер паспорта: " + t.Rows[0][2].ToString();
                    MessageBox.Show(message, "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();
                }
            }
        }
    }
}
