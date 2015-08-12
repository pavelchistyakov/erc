
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;

using System.Windows.Forms;

namespace ERC
{
    class SQL
    {
        static string sConn = "Server=MYPC\\ADK;Database=UtilitiesAccounting_DB;Trusted_Connection=True;";
        static SqlConnection conn;
        static SQL()
        {
            conn = new SqlConnection(sConn);
        }

        public static DataTable FillTable(string sSelect)
        {
            SqlDataAdapter da = new SqlDataAdapter(sSelect, conn);
            DataTable t = new DataTable();
            da.Fill(t);

            return t;
        }

        public static DataTable FillTable(string sSelect, out SqlDataAdapter da)
        {
            da = new SqlDataAdapter(sSelect, conn);
            SqlCommandBuilder b = new SqlCommandBuilder(da);
            //da.InsertCommand.CommandText;
            //da.DeleteCommand.CommandText;
            //da.UpdateCommand.CommandText;

            DataTable t = new DataTable();

            try
            {
                da.Fill(t);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message + "\r\n");
                return null;
            }

            return t;
        }

        public static void SaveTable(SqlDataAdapter da, DataTable t)
        {
            try
            {
                da.Update(t);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message + "\r\n");
                return;
            }
            catch (DBConcurrencyException ex)
            {
                MessageBox.Show(ex.Message + "\r\n");
                return;
            }
        }

        public static bool ExecuteSQL(string sSQL)
        {
            SqlCommand cmd = new SqlCommand(sSQL, conn);
            conn.Open();

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message + "\r\n");
                return false;
            }
            catch (DBConcurrencyException ex)
            {
                MessageBox.Show(ex.Message + "\r\n");
                return false;
            }
            finally
            {
                conn.Close();
            }
            return true;
        }
    }
}
