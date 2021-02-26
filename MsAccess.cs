using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReopenIssueToAccess
{
    class MsAccess
    {
        private OleDbConnection cn;
        private string strConnection;
        private static MsAccess Instance;
        public string FilePath
        {
            private get;
            set;
        }



        private MsAccess()
        {
            FilePath = @"E:\02ATD\DB\NodesDatabase.accdb";
            cn = null;
        }


        public static MsAccess GetInstance()
        {
            if (null == Instance)
            {
                Instance = new MsAccess();
            }
            return Instance;
        }

        public bool Open()
        {
            bool bSucc = true;
            strConnection = @"Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + FilePath ;
            try
            {
                cn = new OleDbConnection(strConnection);
                cn.Open();
            }
            catch (OleDbException ex)
            {
                bSucc = false;
            }
            catch (Exception ex)
            {
                bSucc = false;
            }
            return bSucc;
        }

        public bool Close()
        {
            bool bSucc = true;
            try
            {
                if (null == cn)
                {
                    cn.Close();
                    cn.Dispose();
                }
            }
            catch (OleDbException ex)
            {
                bSucc = false;
            }
            catch (Exception ex)
            {
                bSucc = false;
            }
            return bSucc;
        }

        public bool ExecuteSQL(string strSQL, ref DataSet ds)
        {
            if (null == ds)
            {
                ds = new DataSet();

            }

            bool bSucc = true;
            try
            {
                OleDbDataAdapter oleAdapter = new OleDbDataAdapter(strSQL, cn);
                oleAdapter.Fill(ds);
                oleAdapter.Dispose();
            }
            catch (OleDbException ex)
            {
                bSucc = false;
            }
            catch (Exception ex)
            {
                bSucc = false;
            }
            return bSucc;
        }

        public bool ExecuteSQL(string strSQL, ref DataTable dt)
        {
            if (null == dt)
            {
                dt = new DataTable();

            }

            bool bSucc = true;
            try
            {
                OleDbDataAdapter oleAdapter = new OleDbDataAdapter(strSQL, cn);
                oleAdapter.Fill(dt);
                oleAdapter.Dispose();
            }
            catch (OleDbException ex)
            {
                bSucc = false;
            }
            catch (Exception ex)
            {
                bSucc = false;
            }
            return bSucc;
        }

        public bool ExecuteSQL(string strSQL)
        {
            bool bSucc = true;
            try
            {
                OleDbCommand cmd = new OleDbCommand(strSQL, cn);
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                bSucc = false;
            }
            catch (Exception ex)
            {
                bSucc = false;
            }
            return bSucc;
        }
    }
}
