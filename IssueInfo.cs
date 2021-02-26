using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReopenIssueToAccess
{
    class IssueInfo
    {
        public string DBFilePath
        {
            private get;
            set;
        }
        public DataTable GetIssueSet(string strProjectName, string strIssueType)
        {
            MsAccess.GetInstance().FilePath = this.DBFilePath;
            MsAccess.GetInstance().Open();
            DataTable dt = null;
            string strSQL = "SELECT Key, IssueType, Status,Priority, Resolution, Created, Resolved FROM " + strProjectName + " WHERE IssueType='" + strIssueType + "'";
            MsAccess.GetInstance().ExecuteSQL(strSQL, ref dt);
            MsAccess.GetInstance().Close();
            return dt;
        }
        public DataTable GetIssueSet(string strProjectName)
        {
            MsAccess.GetInstance().FilePath = this.DBFilePath;
            MsAccess.GetInstance().Open();
            DataTable dt = null;
            string strSQL = "SELECT Key, IssueType, Status,Priority, Resolution, Created ,Resolved,DescriptionSize, CommentNum, CommentSize, CommentrNo FROM " + strProjectName;
            MsAccess.GetInstance().ExecuteSQL(strSQL, ref dt);
            MsAccess.GetInstance().Close();
            return dt;
        }
        public void AddToDataSet(string strProjectName,ArrayList arrbugid)
        {
            MsAccess.GetInstance().FilePath = this.DBFilePath;
            MsAccess.GetInstance().Open();
            string strSQL1 = "ALTER table " + strProjectName + " add isReopened INT ";
            MsAccess.GetInstance().ExecuteSQL(strSQL1);
            for (int i = 0; i < arrbugid.Count; i++) {
                string strSQL2 = "UPDATE " + strProjectName + " SET isReopened = 1 WHERE Key = '"+arrbugid[i]+"'";
                MsAccess.GetInstance().ExecuteSQL(strSQL2);
            }
            MsAccess.GetInstance().Close();
        }
    }
}
