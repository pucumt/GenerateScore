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
        private XSSFWorkbook newwb = null;
        private DataTable dt;
        private XSSFSheet sheet1;
        private XSSFSheet newst;
        private int insertIndex = 1;
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
            newwb = new XSSFWorkbook();
            for (var i=0;i< hssfworkbook.NumberOfSheets;i++)
            {
                // calculate average scores of each sheet
                calculateSheet(i);
            }
            FileStream file = new FileStream(@"test.xlsx", FileMode.Create);
            newwb.Write(file);
            file.Close();

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
            // 新的sheet
            newst = newwb.CreateSheet(sheet1.SheetName) as XSSFSheet;
            // 课程名  年级  老师  手机号 考试名称 优等人数 良等人数 有效人数
            newst.GetRow(0).GetCell(0).SetCellValue("课程名");
            newst.GetRow(0).GetCell(1).SetCellValue("年级");
            newst.GetRow(0).GetCell(2).SetCellValue("老师");
            newst.GetRow(0).GetCell(3).SetCellValue("手机号");
            newst.GetRow(0).GetCell(4).SetCellValue("考试名称");
            newst.GetRow(0).GetCell(5).SetCellValue("优等人数");
            newst.GetRow(0).GetCell(6).SetCellValue("良等人数");
            newst.GetRow(0).GetCell(7).SetCellValue("有效人数");
            insertIndex = 1;
            int m = 7;
            var testCell = sheet1.GetRow(0).GetCell(m);
            while (testCell != null)
            {   // 测试分别运算出结果
                calculateTest(m, testCell.StringCellValue);
                m++;
                testCell = sheet1.GetRow(0).GetCell(m);
            }
        }

        private void calculateTest(int m, string testName)
        {
            setDataTable(m);

            DataView dataView = dt.DefaultView;
            DataTable dataTableDistinct = dataView.ToTable(true, "grade");

            for (var j = 0; j < dataTableDistinct.Rows.Count; j++)
            {
                // 按照年级排名
                calculateGrade(testName, dataTableDistinct.Rows[j]["grade"].ToString());
            }
        }

        private void calculateGrade(string testName, string gradeTxt)
        {
            DataView dv = dt.DefaultView;
            dv.RowFilter = "grade = '" + gradeTxt + "' and score> " + cmbScore.Text;
            dv.Sort = "score desc";
            DataTable dtLast = dv.ToTable(false, "score");
            var total = dtLast.Rows.Count;
            int p30 = (int)Math.Ceiling(total * 0.3);
            int p60 = (int)Math.Ceiling(total * 0.6);

            int p30Score = (int)dtLast.Rows[p30]["score"]; // 优的分数
            int p60Score = (int)dtLast.Rows[p60]["score"]; // 良的分数

            // 拉出所有班级信息
            DataView dvClass = dt.DefaultView;
            DataTable dataTableClasses = dvClass.ToTable(true, "className");

            for (var j = 0; j < dataTableClasses.Rows.Count; j++)
            {
                string className = dataTableClasses.Rows[j]["className"].ToString();
                string teacher = dataTableClasses.Rows[j]["teacher"].ToString();
                string mobile = dataTableClasses.Rows[j]["mobile"].ToString();
                newst.GetRow(insertIndex).GetCell(0).SetCellValue(className);
                newst.GetRow(insertIndex).GetCell(1).SetCellValue(gradeTxt);
                newst.GetRow(insertIndex).GetCell(2).SetCellValue(teacher);
                newst.GetRow(insertIndex).GetCell(3).SetCellValue(mobile);
                newst.GetRow(insertIndex).GetCell(4).SetCellValue(testName);
                // 算出优率和良率，然后输出
                DataView dvScore = dt.DefaultView;
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score>= " + p30Score;
                newst.GetRow(insertIndex).GetCell(5).SetCellValue(dvScore.Count);
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score<"+ p30Score + " and score>= " + p60Score;
                newst.GetRow(insertIndex).GetCell(6).SetCellValue(dvScore.Count);
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score>= " + cmbScore.Text;
                newst.GetRow(insertIndex).GetCell(7).SetCellValue(dvScore.Count);
                insertIndex++;
            }
        }

        private void setDataTable(int j)
        {
            // 输入成绩到 table
            dt = new DataTable();
            dt.Columns.Add(new DataColumn("grade"));
            dt.Columns.Add(new DataColumn("score", typeof(float)));
            dt.Columns.Add(new DataColumn("className"));
            dt.Columns.Add(new DataColumn("teacher"));
            dt.Columns.Add(new DataColumn("mobile"));
            var i = 1;
            var s = j;// score column index
            while (sheet1.GetRow(i) != null)
            {
                DataRow row = dt.NewRow();
                row["grade"] = sheet1.GetRow(i).GetCell(3).StringCellValue;
                row["className"] = sheet1.GetRow(i).GetCell(0).StringCellValue;
                row["teacher"] = sheet1.GetRow(i).GetCell(5).StringCellValue;
                row["mobile"] = sheet1.GetRow(i).GetCell(6).StringCellValue;

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
