using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FastReport;
using FastReport.Utils;
using FastReport.Data;
using FastReport.Barcode;
using FastReport.Controls;
using System.Data.SQLite;
using System.Threading;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel; 


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private string filename;
        public Form1()
        {
            InitializeComponent();
            progressBar1.Visible = false;


            // report.PrintPrepared();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                filename=(dlg.FileName);
                Thread workerThread1 = new Thread(new ThreadStart(excel2db));
                workerThread1.Start();

            }
            
            
        }

        void excel2db()
        {
            
            
            try
            {
                this.Invoke(new setStatusDelegate1(setStatus));
                dbHepler db = new dbHepler();
                db.open();
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
                    Excel.Workbook wb = excel.Application.Workbooks._Open(filename);
                    //取得第一个工作薄  
                    Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                   
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
                        this.Invoke(new setStatusDelegate(setStatus), i * 100 / (rowsint - 2));
                        strinsert = String.Format("insert into [20028] (name,XH,ZKZH,级别语言) values ('{0}','{1}','{2}','{3}')", arry3[i, 1].ToString(), arry4[i, 1].ToString(), arry2[i, 1].ToString(), arry1[i, 1].ToString());
                        db.Query(strinsert);
                    }

                    db.close();
                    this.Invoke(new setStatusDelegate1(setStatus2));
                    
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        
        }
        private delegate void setStatusDelegate(int num);
        private delegate void setStatusDelegate1();
        private void setStatus(int num)
        {
            progressBar1.Visible = true;
            progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Blocks;
            progressBar1.Value = num;


        }
        private void setStatus()
        {
            progressBar1.Visible = true;
            progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;

        }
        private void setStatus2()
        {
            progressBar1.Visible = false;
            

        }
     
        private void button3_Click(object sender, EventArgs e)
        {

            dbHepler db = new dbHepler();

            string sql = "select * from [20028]  limit 20";
           // string sql = "select * from [20028] where XH='128590'";
            DataSet Student = db.LoadData(sql);
            Report report = new Report();
            report.Load("Untitled.frx");


            // register the "Products" table

            report.RegisterData(Student.Tables[0], "Student");

            // enable it to use in a report

            report.GetDataSource("Student").Enabled = true;


            GroupHeaderBand group1 = (GroupHeaderBand)report.FindObject("GroupHeader1");



            group1.Height = Units.Centimeters * 1;

            // set group condition

            group1.Condition = "[Student.XH]";



            DataBand data = (DataBand)report.FindObject("data1");
            // register the "Products" table




            data.DataSource = report.GetDataSource("Student");



            TextObject text2 = new TextObject();

            text2.Name = "Text2";

            text2.Bounds = new RectangleF(0, 0,

              Units.Centimeters * 2, Units.Centimeters * 1);

            text2.Text = "[Student.name]";

            text2.Font = new Font("Tahoma", 10, FontStyle.Bold);
            data.Objects.Add(text2);

            // report.SetParameterValue("Parameter", "[Customers.CustLName]");
            PictureObject pic = (PictureObject)report.FindObject("Picture1");
            pic.Image = Image.FromFile("100001.jpg");

         
              // run the report*/
            report.Preview = this.previewControl1;
            report.Prepare();
            report.Show();
        }
    }
}
