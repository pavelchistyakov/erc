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
    public partial class UKReportingForm : Form
    {
        Word._Application oWord;
        private const string TemplatePath = "C:\\Templates\\";
        private const string SavedDocumentsPath = "C:\\Saved Documents\\";
        public UKReportingForm()
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
                        oWord = new Word.Application();
                        Word._Document oDoc = oWord.Documents.Add(TemplatePath + "Отчет_по_управляющим_компаниям.dotx");
                        oDoc.Bookmarks["Месяц"].Range.Text = month.ToString();
                        oDoc.Bookmarks["Год"].Range.Text = year.ToString();
                        t = SQL.FillTable("select * from Управляющая_компания");
                        int i = 2;
                        foreach (DataRow row in t.Rows)
                        {                           
                            
                            int sum = 0;
                            int negative_sum = 0;
                            DataTable flats_id = SQL.FillTable("select ID_квартиры from Квартира where ID_УК = " + row[0].ToString());                            
                            foreach (DataRow flat in flats_id.Rows)
                            {
                                DataTable payments = SQL.FillTable("select Сумма from Оплата where ID_квартиры = " + flat[0].ToString() + " and Месяц=" +
                                    month + " and Год = " + year);
                                foreach (DataRow dr in payments.Rows)
                                {
                                    
                                    sum += Convert.ToInt32(dr[0]);
                                }

                                DataTable debts = SQL.FillTable("select Сумма from Долг where ID_квартиры = " + flat[0].ToString() + " and Месяц=" +
                                    month + " and Год = " + year);
                                foreach (DataRow dr in debts.Rows)
                                {
                                    
                                    negative_sum += Convert.ToInt32(dr[0]);
                                }
                            }
                            oDoc.Tables[1].Rows.Add();
                            oDoc.Tables[1].Cell(i, 1).Range.Text = row[1].ToString();
                            oDoc.Tables[1].Cell(i, 2).Range.Text = sum.ToString();
                            oDoc.Tables[1].Cell(i, 3).Range.Text = negative_sum.ToString();
                            i++;

                            
                        }
                        DateTime dt = DateTime.Now;
                        oDoc.SaveAs(SavedDocumentsPath + "МОУК-" + month.ToString() + "-" + year.ToString() + "-" + dt.Day + dt.Month + dt.Year + dt.Hour + dt.Minute +
                        dt.Second + ".docx");
                        oDoc.Close();
                        MessageBox.Show("Отчёты успешно составлены!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
