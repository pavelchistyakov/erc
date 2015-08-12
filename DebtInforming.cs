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
    public partial class DebtInformingForm : Form
    {
        private const string TemplatePath = "C:\\Templates\\";
        private const string SavedDocumentsPath = "C:\\Saved Documents\\";

        Word._Application oWord;

        public DebtInformingForm()
        {
            InitializeComponent();
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
                foreach(DataRow row_debt in debt_info.Rows)
                {
                    oDoc.Tables[1].Rows.Add();
                    oDoc.Tables[1].Cell(i, 1).Range.Text = row_debt[1].ToString();
                    oDoc.Tables[1].Cell(i, 2).Range.Text = row_debt[0].ToString();
                    oDoc.Tables[1].Cell(i, 3).Range.Text = row_debt[2].ToString();
                    i++;
                }
                DateTime dt = DateTime.Now;
                oDoc.SaveAs(SavedDocumentsPath + "ИОЗ-" + flat_info.Rows[0][1].ToString() + "-" + flat_info.Rows[0][2].ToString() + "-" +
                flat_info.Rows[0][3].ToString()+ "-" + flat_info.Rows[0][4].ToString() + "-" +dt.Day + dt.Month+ dt.Year + dt.Hour + dt.Minute +
                dt.Second + ".docx");
                oDoc.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AllInformingInterface();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        
    }
}
