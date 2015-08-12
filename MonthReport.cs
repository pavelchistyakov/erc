using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ERC
{
    public partial class MonthReportForm : Form
    {
        Word._Application oWord;
        private const string TemplatePath = "C:\\MenshikovaLab\\Templates\\";
        private const string SavedDocumentsPath = "C:\\MenshikovaLab\\Saved Documents\\";
        public MonthReportForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable t;
            int month = 0; int year = 0;
            if (Int32.TryParse(textBox1.Text, out month) == true && Int32.TryParse(textBox2.Text, out year) == true)
            {
                if (month >= 1 && month <= 12)
                {
                    t = SQL.FillTable("select * from Оплата where Месяц = " + month + " and Год = " + year);
                    if (t.Rows.Count == 0) { MessageBox.Show("За данный месяц расчёт не проводился!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                    else
                    {
                        t = SQL.FillTable("select * from Управляющая_компания");
                        foreach(DataRow row in t.Rows)
                        {
                            oWord = new Word.Application();
                            int flats_total = 0;
                            int payments_total = 0;
                            int debts_total = 0;
                            int sum = 0;
                            int negative_sum = 0;
                            DataTable flats_id = SQL.FillTable("select ID_квартиры from Квартира where ID_УК = " + row[0].ToString());
                            flats_total = flats_id.Rows.Count;
                            foreach(DataRow flat in flats_id.Rows)
                            {
                                DataTable payments = SQL.FillTable("select Сумма from Оплата where ID_квартиры = " + flat[0].ToString() + " and Месяц=" +
                                    month + " and Год = " + year);
                                foreach (DataRow dr in payments.Rows)
                                {
                                    if (Convert.ToInt32(dr[0]) > 0) payments_total++;
                                    sum += Convert.ToInt32(dr[0]);
                                }

                                DataTable debts = SQL.FillTable("select Сумма from Долг where ID_квартиры = " + flat[0].ToString() + " and Месяц=" +
                                    month + " and Год = " + year);
                                foreach (DataRow dr in debts.Rows)
                                {
                                    if (Convert.ToInt32(dr[0]) > 0) debts_total++;
                                    negative_sum += Convert.ToInt32(dr[0]);
                                }
                            }

                            Word._Document oDoc = oWord.Documents.Add(TemplatePath + "Месячный_отчёт.dotx");

                            oDoc.Bookmarks["Месяц"].Range.Text = month.ToString();
                            oDoc.Bookmarks["Год"].Range.Text = year.ToString();
                            oDoc.Bookmarks["УК"].Range.Text = row[1].ToString();
                            oDoc.Bookmarks["Квартир"].Range.Text = flats_total.ToString();
                            oDoc.Bookmarks["Оплат"].Range.Text = payments_total.ToString();
                            oDoc.Bookmarks["Долги"].Range.Text = debts_total.ToString();
                            oDoc.Bookmarks["Сумма"].Range.Text = sum.ToString();
                            oDoc.Bookmarks["Недостача"].Range.Text = negative_sum.ToString();

                            DateTime dt = DateTime.Now;
                            oDoc.SaveAs(SavedDocumentsPath + "МО-" + row[1].ToString() + "-" + month.ToString()+ "-" + year.ToString() + "-" + dt.Day + dt.Month + dt.Year + dt.Hour + dt.Minute +
                            dt.Second + ".docx");
                            oDoc.Close();
                            MessageBox.Show("Отчёты успешно составлены!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Значение месяца должно быть от 1 от 12", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Проверьте правильность ввода полей!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
