using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Data;

namespace WindowsFormsApplication1
{
    class dbHepler
    {
        public SQLiteConnection sql_con;
        private  SQLiteTransaction _sqLiteTransaction;
        public SQLiteCommand sql_cmd;
        public SQLiteDataAdapter DB;
        public DataSet DS = new DataSet();
        private DataTable DT = new DataTable();
        private String myPath;
        private String filename = "\\quekaomingdan.csv";

        public void SetConnection()
        {
            sql_con = new SQLiteConnection
                ("Data Source=DemoT.db");
        }
        public void open()
        {
            SetConnection();
            sql_con.Open();
            _sqLiteTransaction = sql_con.BeginTransaction();
            sql_cmd = sql_con.CreateCommand();
        }
        public void close()
        {
            _sqLiteTransaction.Commit();
            _sqLiteTransaction.Dispose(); 
            sql_con.Close();
        }
        public void Query(string txtQuery)
        {
           
            sql_cmd.CommandText = txtQuery;
            sql_cmd.ExecuteNonQuery();
            
        }
        public void ExecuteQuery(string txtQuery)
        {
            SetConnection();
            sql_con.Open();
            sql_cmd = sql_con.CreateCommand();
            sql_cmd.CommandText = txtQuery;
            sql_cmd.ExecuteNonQuery();
            sql_con.Close();
        }

        private void LoadData1(string zkzh)
        {
            SetConnection();
            sql_con.Open();
           
            //sql_cmd = sql_con.CreateCommand();
            SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand();
            cmd.CommandText = "select name, XH,SFZH,ZKZH from [20028] where [ZKZH]=" + zkzh;
            cmd.Connection = sql_con;
            SQLiteDataReader dr;

            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
               


            }
            /*DB = new SQLiteDataAdapter(CommandText,sql_con); 
            DS.Reset(); 
            DB.Fill(DS); 
		
            if (DS.Tables[0].Rows.Count> 0)
            {
                txbName.Text = DS.Tables[0].Rows[0]["name"].ToString();
                txbXH.Text = DS.Tables[0].Rows[0]["XH"].ToString();
                txbZKZH.Text = DS.Tables[0].Rows[0]["ZKZH"].ToString();
                SFZHCurrent = DS.Tables[0].Rows[0]["SFZH"].ToString();

                String[] stu = new String[4];
                stu[0] = txbName.Text;
                stu[1] = txbXH.Text;
                stu[2] = txbZKZH.Text;
                stu[3] = DS.Tables[0].Rows[0]["SFZH"].ToString();

                Stu.Clear();
                Stu.Add(stu);


            }*/
            else
            {
                //MessageBox.Show("未查到该生信息");
            }


            sql_con.Close();
        }
        public DataSet LoadData(string CommandText)
        {
            SetConnection();
            sql_con.Open();
            //sql_cmd = sql_con.CreateCommand();
            SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand();
           // cmd.CommandText = "select name, XH,SFZH,ZKZH from [20028] where [ZKZH]=" + zkzh;
           // cmd.Connection = sql_con;
            
            DB = new SQLiteDataAdapter(CommandText,sql_con); 
            DS.Reset(); 
            DB.Fill(DS); 
		
            


            sql_con.Close();
            return DS;
        
        }

    }
}
