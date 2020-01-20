using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


namespace 标本查看
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            IWorkbook wb = null;
            ISheet ws = null;

            string filepath = FileDialog();
            FileStream fs = null;

            if (filepath == null)
            {
                return;
            }
            else
            {
                fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            }

            //if (filepath.IndexOf(".xlsx") > 0)
            //{
            //    wb = new XSSFWorkbook(fs);
            //}
            //else if (filepath.IndexOf(".xls") > 0)
            //{
            //    wb = new XSSFWorkbook(fs);
            //}
            //else
            //{
            //    return;
            //}

            wb = WorkbookFactory.Create(fs);
            ws = wb.GetSheetAt(0);

            DataTable dt = new DataTable(ws.SheetName);
            dt.Rows.Clear();
            dt.Columns.Clear();

            int startRow = ws.FirstRowNum;
            int rowCount = ws.LastRowNum;

            IRow firstRow = ws.GetRow(startRow);

            #region 默认列数等于第一行的列数
            int colCount = firstRow.LastCellNum;
            for (int c=0; c < colCount; c++)
            {
                if (firstRow.GetCell(c) != null)
                {
                    dt.Columns.Add(firstRow.GetCell(c).StringCellValue);
                }
            }
            #endregion

            #region 添加数据至dt
            for (int r=startRow+1; r<=rowCount; r++)
            {
                IRow row = ws.GetRow(r);
                DataRow dataRow = dt.NewRow();
                
                for (int c=0; c<colCount; c++)
                {
                    if (row.GetCell(c) != null)
                    {
                        dataRow[c] = row.GetCell(c).ToString();
                    }
                }
                dt.Rows.Add(dataRow);
            }
            #endregion

            dataGridView1.DataSource = dt;
        }

        public string FileDialog()
        {
            string filepath = null;

            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = "Excel文件(*.xls; *.xlsx; *.csv)|*.xls; *.xlsx; *.csv";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filepath = openFileDialog1.FileName;
            }
            return filepath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = "";
        }
    }
}
