using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ReopenIssueToAccess
{
    public partial class Form1 : Form
    {
        FolderBrowserDialog folder = new FolderBrowserDialog();
        Microsoft.Office.Interop.Excel.Application app = null;
        int col;
        int row;
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            folder.ShowDialog();
            textBox1.Text = folder.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = dialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            label3.ForeColor = Color.Red;
            label3.Text = "running";

            string rootPath = textBox1.Text;
            string DBfilepath = textBox2.Text;
            int index = -1;
            string subStr = "";
            DirectoryInfo directory = new DirectoryInfo(rootPath);
            foreach (FileInfo file in directory.GetFiles()) {
                ArrayList arrBugID = new ArrayList();
                string fileName = file.FullName;
                index = fileName.LastIndexOf(".");
                subStr = fileName.Substring(index + 1);
                if (subStr != "xlsx")
                {
                    label3.ForeColor = Color.Red;
                    label3.Text = "存在非Excel";
                    return;
                }
                string strProjectName = fileName.Substring(fileName.LastIndexOf("\\") + 1, fileName.Length - fileName.LastIndexOf("\\") - 1);
                strProjectName = strProjectName.Substring(0, strProjectName.IndexOf("_"));
                List<List<object>> InputMatrix;
                int RowNum;
                ReadExcel(fileName, out InputMatrix, out RowNum);
                for (int i = 1; i < RowNum; i++) {
                    arrBugID.Add(InputMatrix[i][1]);
                }
                //将bug写入access
                IssueInfo issueInfo = new IssueInfo();
                issueInfo.DBFilePath = DBfilepath;
                issueInfo.AddToDataSet(strProjectName, arrBugID);
  
                label3.ForeColor = Color.Green;
                label3.Text = "success!";
            }
        }
       
        public void ReadExcel(string FilePath, out List<List<object>> InputMatrix, out int RowNum)
        {
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            app = new Microsoft.Office.Interop.Excel.Application();
            workBook = app.Workbooks.Open(FilePath);
            Worksheet worksheet = (Worksheet)workBook.Worksheets[1];//选择sheet
            col = worksheet.UsedRange.CurrentRegion.Columns.Count;
            row = worksheet.UsedRange.CurrentRegion.Rows.Count;
            InputMatrix = NewListMatrixOfObject(row, col);
            object[,] current;
            current = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row, col]].Value2;

            int k = 1;
            for (int i = 0; i <= row - 1; i++)
            {
                if (current[i + 1, 18] != null)//读取的excel列数
                {
                    for (int j = 1; j <= col; j++)
                    {
                        InputMatrix[k - 1][j - 1] = current[i + 1, j];
                    }
                    k++;
                }
            }
            RowNum = k - 1;
            app.Quit();
            app = null;
            workBook = null;
        }
        public static List<List<object>> NewListMatrixOfObject(int iRowCount, int iColCount)
        {
            List<List<object>> matrix = new List<List<object>>(iRowCount);
            for (int i = 0; i < iRowCount; i++)
            {
                matrix.Add(new List<object>(iColCount));
                for (int j = 0; j < iColCount; j++)
                {
                    matrix[i].Add("");
                }
            }
            return matrix;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
