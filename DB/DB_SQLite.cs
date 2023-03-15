using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 簡易倉儲系統.DB
{
    internal class DB_SQLite
    {
        //https://blog.poychang.net/note-sqlite/

        /// <summary>建立資料庫連線</summary>
        /// <param name="database">資料庫名稱</param>
        /// <returns></returns>
        public SQLiteConnection OpenConnection(string database)
        {
            var conntion = new SQLiteConnection()
            {
                ConnectionString = $"Data Source={database};Version=3;New=False;Compress=True;"
            };
            if (conntion.State == ConnectionState.Open) conntion.Close();
            conntion.Open();
            return conntion;
        }

        /// <summary>建立新資料庫</summary>
        /// <param name="database">資料庫名稱</param>
        public void CreateDatabase(string database)
        {
            var connection = new SQLiteConnection()
            {
                ConnectionString = $"Data Source={database};Version=3;New=True;Compress=True;"
            };
            connection.Open();
            connection.Close();
        }

        /// <summary>建立新資料表</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlCreateTable">建立資料表的 SQL 語句</param>
        public void CreateTable(string database, string sqlCreateTable)
        {
            var connection = OpenConnection(database);
            //connection.Open();
            var command = new SQLiteCommand(sqlCreateTable, connection);
            var mySqlTransaction = connection.BeginTransaction();
            try
            {
                command.Transaction = mySqlTransaction;
                command.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (Exception ex)
            {
                mySqlTransaction.Rollback();
                throw (ex);
            }
            if (connection.State == ConnectionState.Open) connection.Close();
        }

        /// <summary>新增\修改\刪除資料</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlManipulate">資料操作的 SQL 語句</param>
        public void Manipulate(string database, string sqlManipulate)
        {
            var connection = OpenConnection(database);
            var command = new SQLiteCommand(sqlManipulate, connection);
            var mySqlTransaction = connection.BeginTransaction();
            try
            {
                command.Transaction = mySqlTransaction;
                command.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (Exception ex)
            {
                mySqlTransaction.Rollback();
                throw (ex);
            }
            if (connection.State == ConnectionState.Open) connection.Close();
        }

        /// <summary>讀取資料</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlQuery">資料查詢的 SQL 語句</param>
        /// <returns></returns>
        public DataTable GetDataTable(string database, string sqlQuery)
        {
            var connection = OpenConnection(database);
            var dataAdapter = new SQLiteDataAdapter(sqlQuery, connection);
            var myDataTable = new DataTable();
            var myDataSet = new DataSet();
            myDataSet.Clear();
            dataAdapter.Fill(myDataSet);
            myDataTable = myDataSet.Tables[0];
            if (connection.State == ConnectionState.Open) connection.Close();
            return myDataTable;
        }

        /// <summary>
        /// Datatable轉出Datagridview
        /// 全部刪除重繪
        /// </summary>
        /// <param name="DT"></param>
        /// <param name="DGV"></param>
        /// <returns></returns>
        public static Boolean DatatableToDatagridview(DataTable DT, DataGridView DGV)
        {
            try
            {
                DGV.Rows.Clear();
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataGridViewRow _data = new DataGridViewRow();
                    _data.CreateCells(DGV);
                    List<string> strings = new List<string>();
                    for (int j = 0; j < DT.Columns.Count; j++)
                    {
                        string _value = DT.Rows[i][j].ToString();
                        strings.Insert(j, _value);
                    }
                    _data.SetValues(strings.ToArray());
                    DGV.Rows.Insert(0, _data);
                    DGV.Rows[0].Selected = true;
                    DGV.CurrentCell = DGV.Rows[0].Cells[0];
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
