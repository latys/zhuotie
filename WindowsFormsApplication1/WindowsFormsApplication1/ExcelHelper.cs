using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;  

namespace WindowsFormsApplication1
{
    class ExcelHelper
    {
        public void OpenExcel(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application  
            if (excel == null)
            {
                //this.label1.Text = "Can't access excel";
            }
            else
            {
                excel.Visible = false;
                excel.UserControl = true;
                // 以只读的形式打开EXCEL文件  
                Excel.Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄  
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列)  
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数  
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数  
                //取得数据范围区域   (不包括标题列)    
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                object[,] arry1 = (object[,])rng1.Value2;   //get range's value  
                object[,] arry2 = (object[,])rng2.Value2;
                object[,] arry3 = (object[,])rng3.Value2;   //get range's value  
                object[,] arry4 = (object[,])rng4.Value2;
                //将新值赋给一个数组  
                string[,] arry = new string[rowsint - 1, 4];
                //for (int i = 1; i <= rowsint - 1; i++)  
                for (int i = 1; i <= rowsint - 2; i++)
                {

                    arry[i - 1, 0] = arry1[i, 1].ToString();

                    arry[i - 1, 1] = arry2[i, 1].ToString();

                    arry[i - 1, 2] = arry3[i, 1].ToString();

                    arry[i - 1, 3] = arry4[i, 1].ToString();
                }
                string a = "";
                for (int i = 0; i <= rowsint - 3; i++)
                {
                    a += arry[i, 0] + "|" + arry[i, 1] + "|" + arry[i, 2] + "|" + arry[i, 3] + "\n";

                }
               // this.label1.Text = a;
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程  
            }
            GC.Collect();
        }

        public void excel2db(string strFileName)
        {
             
            dbHepler db = new dbHepler();
            string strdel = "delete  from [20028]";
            string strinsert;
            db.ExecuteQuery(strdel);
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application  
            if (excel == null)
            {
                //this.label1.Text = "Can't access excel";
            }
            else
            {
                excel.Visible = false;
                excel.UserControl = true;
                // 以只读的形式打开EXCEL文件  
                Excel.Workbook wb = excel.Application.Workbooks._Open(strFileName);
                //取得第一个工作薄  
                Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
               /* if (wb.Worksheets.Count > 0)
                {
                    ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                }
                else
                {
                    wb.Worksheets.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1);
                }*/
                //取得总记录行数    (包括标题列)  
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数  
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数  
                //取得数据范围区域   (不包括标题列)    
                Excel.Range rng1 = ws.Cells.get_Range("B2", "B" + rowsint);  //语言级别
                Excel.Range rng2 = ws.Cells.get_Range("G2", "G" + rowsint);  //准考证
                Excel.Range rng3 = ws.Cells.get_Range("H2", "H" + rowsint);  //姓名
                Excel.Range rng4 = ws.Cells.get_Range("W2", "W" + rowsint);  //学号
                object[,] arry1 = (object[,])rng1.Value2;   //get range's value  
                object[,] arry2 = (object[,])rng2.Value2;
                object[,] arry3 = (object[,])rng3.Value2;   //get range's value  
                object[,] arry4 = (object[,])rng4.Value2;
                //将新值赋给一个数组  
                

                string[,] arry = new string[rowsint - 1, 4];
                //for (int i = 1; i <= rowsint - 1; i++)  
               
                for (int i = 1; i <= rowsint - 2; i++)
                {

                    //Form1.progressBar1.Value = i * 100 / (rowsint - 2);
                    
                    strinsert = String.Format("insert into [20028] (name,XH,ZKZH,级别语言) values ('{0}','{1}','{2}','{3}')", arry3[i, 1].ToString(), arry4[i, 1].ToString(), arry2[i, 1].ToString(), arry1[i, 1].ToString());
                    db.ExecuteQuery(strinsert);
                }

               
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程  
            }
            GC.Collect();
            MessageBox.Show("导入成功");
        
        }

    }
}
