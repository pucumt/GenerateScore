using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data.OleDb;
using System.Threading;
using System.Data.SqlClient;

namespace GenerateScore
{
    public partial class Frm_Generator : Form
    {
        private XSSFWorkbook hssfworkbook = null;
        private DataTable dt;
        private XSSFSheet sheet1;
        public Frm_Generator()
        {
            InitializeComponent();
            cmbScore.SelectedIndex = 0;
        }

        private void btn_select_Click(object sender, EventArgs e)
        {
            var result = ofD_file.ShowDialog();
            txtFile.Text = ofD_file.FileName;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {

        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(txtFile.Text.Trim()))
            {
                return;
            }

            try
            {
                using (FileStream fs = File.Open(@txtFile.Text, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
                {
                    //把xls文件读入workbook变量里，之后就可以关闭了  
                    hssfworkbook = new XSSFWorkbook(fs);
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error:"+ex.Message);
            }

            // 获取sheetName
            cmbSheet.Items.Clear();
            cmbSheet.SelectedIndex = -1;
            cmbSheet.Text = "";
            for (var i=0;i< hssfworkbook.NumberOfSheets;i++)
            {
                // calculate average scores of each sheet
                // cmbSheet.Items.Add(hssfworkbook.GetSheetName(i));
                calculateSheet(i);
            }

            // 跳转tab
            tabControl1.SelectTab(1);
            tabPage2.Show();
        }

        private void calculateSheet(int i)
        {
            if (hssfworkbook == null)
            {
                return;
            }

            sheet1 = hssfworkbook.GetSheetAt(i) as XSSFSheet;
            if (sheet1 == null)
            {
                return;
            }

            setDataTable(1);

            DataView dataView = dt.DefaultView;
            DataTable dataTableDistinct = dataView.ToTable(true, "grade");

            for (var j = 0; j < dataTableDistinct.Rows.Count; j++)
            {
               // 按照年级排名

            }
        }

        private void calculateGrade()
        {

        }

        private void cmbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblScore.Text = "";
            if (hssfworkbook==null)
            {
                return;
            }

            sheet1 = hssfworkbook.GetSheetAt(0) as XSSFSheet;

            if(sheet1 == null)
            {
                return;
            }

            // grade
            cmbGrade.Items.Clear();
            cmbGrade.Text = "";
            cmbGrade.SelectedIndex = -1;

            setDataTable();

            DataView dataView = dt.DefaultView;
            DataTable dataTableDistinct = dataView.ToTable(true, "grade");

            for(var j=0;j< dataTableDistinct.Rows.Count;j++)
            {
                cmbGrade.Items.Add(dataTableDistinct.Rows[j]["grade"]);
            }

            // test name
            cmbTest.Items.Clear();
            cmbTest.Text = "";
            cmbTest.SelectedIndex = -1;

            var m = 4;
            var testCell = sheet1.GetRow(0).GetCell(m);
            while(testCell != null)
            {
                cmbTest.Items.Add(testCell.StringCellValue);
                m++;
                testCell = sheet1.GetRow(0).GetCell(m);
            }
            cmbTest.SelectedIndex = 0;
        }

        private void cmbGrade_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cmbGrade.SelectedIndex == -1 || cmbTest.SelectedIndex == -1 || cmbSheet.SelectedIndex == -1)
            {
                return;
            }
            lblScore.Text = "";
            DataView dv = dt.DefaultView;
            dv.RowFilter = "grade = '" + cmbGrade.Text + "' and score> "+cmbScore.Text;
            dv.Sort = "score desc";
            DataTable dtLast = dv.ToTable(false, "score");
            var total = dtLast.Rows.Count;
            int p30 = (int)Math.Ceiling(total * 0.3);
            int p60 = (int)Math.Ceiling(total * 0.6);

            lblScore.Text = "30%: "+ dtLast.Rows[p30]["score"] +"\r\n60%: " + dtLast.Rows[p60]["score"];
        }

        private void cmbScore_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void cmbTest_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void setDataTable(int j)
        {
            // 输入成绩到 table
            dt = new DataTable();
            dt.Columns.Add(new DataColumn("grade"));
            dt.Columns.Add(new DataColumn("score", typeof(float)));
            dt.Columns.Add(new DataColumn("className"));
            var i = 1;
            var s = 6 + j;// score column index
            while (sheet1.GetRow(i) != null)
            {
                DataRow row = dt.NewRow();
                row["grade"] = sheet1.GetRow(i).GetCell(3).StringCellValue;
                row["className"] = sheet1.GetRow(i).GetCell(0).StringCellValue;
                var cell = sheet1.GetRow(i).GetCell(s);
                if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    row["score"] = sheet1.GetRow(i).GetCell(s).NumericCellValue;
                }
                else
                {
                    row["score"] = float.Parse(sheet1.GetRow(i).GetCell(s).StringCellValue);
                }
                
                dt.Rows.Add(row);
                i++;
            }
        }
    }
}
