using System;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using System.Data;

namespace window3
{
    class SQLRequests
    {
        SQLiteConnection m_dbConn = new SQLiteConnection();
        SQLiteCommand m_sqlCmd = new SQLiteCommand();
        public SQLRequests()
        {
        }
        public void Connect(string dbFileName)
        {
            try {
                m_dbConn.Close();
            }
            catch (System.NullReferenceException e){ }
            if (!File.Exists(dbFileName))
                SQLiteConnection.CreateFile(dbFileName);

            try
            {
                m_dbConn = new SQLiteConnection("Data Source=" + dbFileName + ";Version=3;");
                m_dbConn.Open();
                m_sqlCmd.Connection = m_dbConn;
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        public DataTable getResTable(string Query)
        {
            DataTable dTable = new DataTable();
            try
            {
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(Query, m_dbConn);
                adapter.Fill(dTable);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            return dTable;
        }
        public String getRes(string Query)
        {
            DataTable dTable = new DataTable();
            try
            {
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(Query, m_dbConn);
                adapter.Fill(dTable);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            return dTable.Rows[0].ItemArray[0].ToString();
        }
        public bool isThereRes(string Query)
        {
            DataTable dTable = new DataTable();
            try
            {
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(Query, m_dbConn);
                adapter.Fill(dTable);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            return dTable.Rows.Count > 0 ? true : false;
        }
        public void makeQuery(string Query)
        {
            try
            {
                m_sqlCmd.CommandText = Query;
                m_sqlCmd.ExecuteNonQuery();
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        public void Close()
        {
            try
            {
                m_dbConn.Close();
            }
            catch (System.NullReferenceException e) { }
        }
    }
}
