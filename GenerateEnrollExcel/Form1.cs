using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        private DataTable dtTotal;
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
            setTotalTable();
            for (var i=0;i< hssfworkbook.NumberOfSheets;i++)
            {
                // calculate average scores of each sheet
                calculateSheet(i);
            }

            renderTotalSheet(); // 计算出单个老师的优良率

            FileStream file = new FileStream(@"test.xlsx", FileMode.Create);
            newwb.Write(file);
            file.Close();
            MessageBox.Show("处理完毕，请查看test.xlsx");
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
            XSSFRow newRow = newst.CreateRow(0) as XSSFRow;
            // 课程名  年级  老师  手机号 考试名称 优等人数 良等人数 有效人数
            newRow.CreateCell(0).SetCellValue("课程名");
            newRow.CreateCell(1).SetCellValue("年级");
            newRow.CreateCell(2).SetCellValue("老师");
            newRow.CreateCell(3).SetCellValue("手机号");
            newRow.CreateCell(4).SetCellValue("考试名称");
            newRow.CreateCell(5).SetCellValue("优等人数");
            newRow.CreateCell(6).SetCellValue("优等比例");
            newRow.CreateCell(7).SetCellValue("良等人数");
            newRow.CreateCell(8).SetCellValue("良等比例");
            newRow.CreateCell(9).SetCellValue("差等人数");
            newRow.CreateCell(10).SetCellValue("差等比例");
            newRow.CreateCell(11).SetCellValue("有效人数");
            insertIndex = 1;
            int m = 7;
            var testCell = sheet1.GetRow(0).GetCell(m);
            while (testCell != null)
            {   // 测试分别运算出结果
                calculateTest(m, testCell.StringCellValue, sheet1.GetRow(0).GetCell(m+1)==null);
                m++;
                testCell = sheet1.GetRow(0).GetCell(m);
            }
        }

        private void calculateTest(int m, string testName, bool isLast)
        {
            setDataTable(m);

            DataView dataView = dt.DefaultView;
            DataTable dataTableDistinct = dataView.ToTable(true, "grade");

            for (var j = 0; j < dataTableDistinct.Rows.Count; j++)
            {
                // 按照年级排名
                calculateGrade(testName, dataTableDistinct.Rows[j]["grade"].ToString(), isLast);
            }
        }

        private void calculateGrade(string testName, string gradeTxt, bool isLast)
        {
            DataView dv = dt.DefaultView;
            dv.RowFilter = "grade = '" + gradeTxt + "' and score> " + cmbScore.Text;
            dv.Sort = "score desc";
            DataTable dtLast = dv.ToTable(false, "score");
            var total = dtLast.Rows.Count;
            if(total==0)
            {
                return; 
            }
            int p30 = (int)Math.Ceiling(total * 0.3);
            int p60 = (int)Math.Ceiling(total * 0.6);

            float p30Score = (float)dtLast.Rows[p30]["score"]; // 优的分数
            float p60Score = (float)dtLast.Rows[p60]["score"]; // 良的分数

            // 拉出所有班级信息
            DataView dvClass = dt.DefaultView;
            DataTable dataTableClasses = dvClass.ToTable(true, "className", "teacher", "mobile");

            for (var j = 0; j < dataTableClasses.Rows.Count; j++)
            {
                string className = dataTableClasses.Rows[j]["className"].ToString();
                string teacher = dataTableClasses.Rows[j]["teacher"].ToString();
                string mobile = dataTableClasses.Rows[j]["mobile"].ToString();
                XSSFRow newRow = newst.CreateRow(insertIndex) as XSSFRow;
                newRow.CreateCell(0).SetCellValue(className);
                newRow.CreateCell(1).SetCellValue(gradeTxt);
                newRow.CreateCell(2).SetCellValue(teacher);
                newRow.CreateCell(3).SetCellValue(mobile);
                newRow.CreateCell(4).SetCellValue(testName);

                // 算出优率和良率，然后输出
                DataView dvScore = dt.DefaultView;
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score>= " + p30Score;
                int p30Count = dvScore.Count;
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score<" + p30Score + " and score>= " + p60Score;
                int p60Count = dvScore.Count;
                dvScore.RowFilter = "grade = '" + gradeTxt + "' and className='" + className + "' and score>= " + cmbScore.Text;
                int pTotal = dvScore.Count;
                int leftCount = pTotal - p30Count - p60Count;

                newRow.CreateCell(5).SetCellValue(p30Count);
                newRow.CreateCell(6).SetCellValue(Math.Round((double)p30Count/ pTotal,2));
                newRow.CreateCell(7).SetCellValue(p60Count);
                newRow.CreateCell(8).SetCellValue(Math.Round((double)p60Count / pTotal, 2));
                newRow.CreateCell(9).SetCellValue(leftCount);
                newRow.CreateCell(10).SetCellValue(Math.Round((double)leftCount / pTotal, 2));
                newRow.CreateCell(11).SetCellValue(pTotal);

                if(isLast)
                {
                    // set data to calculate single teacher
                    DataRow row = dtTotal.NewRow();
                    row["teacher"] = teacher;
                    row["mobile"] = mobile;
                    row["p30"] = p30Count;
                    row["p60"] = p60Count;
                    row["pLast"] = leftCount;
                    row["pTotal"] = pTotal;
                    dtTotal.Rows.Add(row);
                }

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

        private void setTotalTable()
        {// 总的分数算法
            dtTotal = new DataTable();
            dtTotal.Columns.Add(new DataColumn("teacher"));
            dtTotal.Columns.Add(new DataColumn("mobile"));
            dtTotal.Columns.Add(new DataColumn("p30", typeof(float)));
            dtTotal.Columns.Add(new DataColumn("p60", typeof(float)));
            dtTotal.Columns.Add(new DataColumn("pLast", typeof(float)));
            dtTotal.Columns.Add(new DataColumn("pTotal", typeof(float)));
        }

        private void renderTotalSheet()
        {
            XSSFSheet totalST = newwb.CreateSheet("汇总") as XSSFSheet;
            XSSFRow newRow = totalST.CreateRow(0) as XSSFRow;
            // 老师  手机号 优等人数 良等人数 有效人数
            newRow.CreateCell(0).SetCellValue("老师");
            newRow.CreateCell(1).SetCellValue("手机号");
            newRow.CreateCell(2).SetCellValue("优等人数");
            newRow.CreateCell(3).SetCellValue("优等比例");
            newRow.CreateCell(4).SetCellValue("良等人数");
            newRow.CreateCell(5).SetCellValue("良等比例");
            newRow.CreateCell(6).SetCellValue("差等人数");
            newRow.CreateCell(7).SetCellValue("差等比例");
            newRow.CreateCell(8).SetCellValue("有效人数");

            DataView dvClass = dtTotal.DefaultView;
            DataTable dataTableClasses = dvClass.ToTable(true, "teacher", "mobile");
            for (var j = 0; j < dataTableClasses.Rows.Count; j++)
            {
                newRow = totalST.CreateRow(j+1) as XSSFRow;
                string teacher = dataTableClasses.Rows[j]["teacher"].ToString();
                string mobile = dataTableClasses.Rows[j]["mobile"].ToString();
                // 老师  手机号 优等人数 良等人数 有效人数
                newRow.CreateCell(0).SetCellValue(teacher);
                newRow.CreateCell(1).SetCellValue(mobile);

                dvClass.RowFilter = "teacher='" + teacher + "' and mobile='" + mobile+"'";
                DataTable newTable = dvClass.ToTable(false, "p30", "p60", "pLast", "pTotal");
                int p30Total = int.Parse(newTable.Compute("sum(p30)", "").ToString());
                int p60Total = int.Parse(newTable.Compute("sum(p60)", "").ToString());
                int pLastTotal = int.Parse(newTable.Compute("sum(pLast)", "").ToString());
                int pTotalTotal = int.Parse(newTable.Compute("sum(pTotal)", "").ToString());

                newRow.CreateCell(2).SetCellValue(p30Total);
                newRow.CreateCell(3).SetCellValue(Math.Round((double)p30Total / pTotalTotal, 2));
                newRow.CreateCell(4).SetCellValue(p60Total);
                newRow.CreateCell(5).SetCellValue(Math.Round((double)p60Total / pTotalTotal, 2));
                newRow.CreateCell(6).SetCellValue(pLastTotal);
                newRow.CreateCell(7).SetCellValue(Math.Round((double)pLastTotal / pTotalTotal, 2));
                newRow.CreateCell(8).SetCellValue(pTotalTotal);
            }
        }
    }
}
