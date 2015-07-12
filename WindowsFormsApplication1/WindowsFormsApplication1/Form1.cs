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
using System.IO;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private string filename;
        Report report = new Report();
        public Form1()
        {
            InitializeComponent();
            progressBar1.Visible = false;
            //初始化考场
            string sql = "SELECT (substr(ZKZH,6,1)||'-'||substr(ZKZH,11,3)) as KAOCHANG from [20028]";

            dbHepler db = new dbHepler();
            DataSet KAOCHANG = db.LoadData(sql);
            comboBox1.DisplayMember = "KAOCHANG";
            comboBox1.ValueMember = "KAOCHANG";
            comboBox1.DataSource = KAOCHANG.Tables[0];
            
            string sql1 = generatesql();
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
                
                string strdel = "delete  from [20028]";
                string strinsert;
                db.ExecuteQuery(strdel);
                db.open();
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
        private delegate string getsql();

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
            Thread workerThread2 = new Thread(new ThreadStart(reportShow));
            workerThread2.Start();
           // reportShow();
        }
        private void reportShow1()
        {
            report.Preview = this.previewControl1;
            report.Prepare();
            report.Show();
        }
        private void reportShow()
        {
            this.Invoke(new setStatusDelegate1(setStatus));
            int i = 0, j = 0;
            dbHepler db = new dbHepler();

            string sql = this.Invoke(new getsql(generatesql)) as string;
            MessageBox.Show(sql);
            // string sql = "select * from [20028] where XH='128590'";
            DataSet Student = db.LoadData(sql);
            report.Pages.Clear();
            //report.Load("Untitled.frx");
            ;

            //DataBand data = (DataBand)report.FindObject("data1");

            for (i = 0; i < Student.Tables[0].Rows.Count / 10; i++)
            {
                ReportPage page1 = new ReportPage();

                report.Pages.Add(page1);

                DataBand data = new DataBand();
                page1.Bands.Add(data);
                for (j = 0; j < 10; j++)
                {


                    TextObject text1 = new TextObject();
                    if (j % 2 == 0)
                    {
                        text1.Bounds = new RectangleF(Units.Centimeters * 4, Units.Centimeters * 3 * j, Units.Centimeters * 5, Units.Centimeters * 0.6f);
                    }
                    else
                    {
                        text1.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 4, Units.Centimeters * 3 * (j - 1), Units.Centimeters * 5, Units.Centimeters * 0.6f);

                    }

                    text1.Text = Student.Tables[0].Rows[10 * i + j]["name"].ToString();
                    data.Objects.Add(text1);


                    PictureObject pic = new PictureObject();
                    if (j % 2 == 0)
                    {
                        pic.Bounds = new RectangleF(0, Units.Centimeters * 3 * j, Units.Centimeters * 2, Units.Centimeters * 2);
                    }
                    else
                    {
                        pic.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 0, Units.Centimeters * 3 * (j - 1), Units.Centimeters * 2, Units.Centimeters * 2);

                    }
                    if (File.Exists("./" + "20"+Student.Tables[0].Rows[10 * i + j]["XH"].ToString().Substring(0,2) + "/" + Student.Tables[0].Rows[10 * i + j]["XH"].ToString()+ ".jpg"))
                        pic.Image = Image.FromFile("./" +"20"+ Student.Tables[0].Rows[10 * i + j]["XH"].ToString().Substring(0,2) + "/" + Student.Tables[0].Rows[10 * i + j]["XH"].ToString() + ".jpg");
                    data.Objects.Add(pic);

                    BarcodeObject bar = new BarcodeObject();
                    if (j % 2 == 0)
                    {
                        bar.Bounds = new RectangleF(0, Units.Centimeters * 3 * j + Units.Centimeters * 2.5f, Units.Centimeters * 8, Units.Centimeters * 2);
                    }
                    else
                    {
                        bar.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 0, Units.Centimeters * 3 * (j - 1) + Units.Centimeters * 2.5f, Units.Centimeters * 8, Units.Centimeters * 2);

                    }
                    //pic.Image = Image.FromFile("100001.jpg");
                    bar.Text = Student.Tables[0].Rows[10 * i + j]["ZKZH"].ToString();
                    data.Objects.Add(bar);

                }

            }

            if (Student.Tables[0].Rows.Count % 10 != 0)
            {
                Console.WriteLine((10 * i + j).ToString());
                ReportPage page2 = new ReportPage();

                report.Pages.Add(page2);

                DataBand data2 = new DataBand();
                page2.Bands.Add(data2);

                for (int k = 10 * (i - 1) + j; k < Student.Tables[0].Rows.Count; k++)
                {


                    TextObject text1 = new TextObject();
                    if (k % 2 == 0)
                    {
                        text1.Bounds = new RectangleF(Units.Centimeters * 4, Units.Centimeters * 3 * (k % 10), Units.Centimeters * 5, Units.Centimeters * 0.6f);
                    }
                    else
                    {
                        text1.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 4, Units.Centimeters * 3 * (k % 10 - 1), Units.Centimeters * 5, Units.Centimeters * 0.6f);

                    }

                    text1.Text = Student.Tables[0].Rows[k]["name"].ToString();
                    data2.Objects.Add(text1);


                    PictureObject pic = new PictureObject();
                    if (k % 2 == 0)
                    {
                        pic.Bounds = new RectangleF(0, Units.Centimeters * 3 * k % 10, Units.Centimeters * 2, Units.Centimeters * 2);
                    }
                    else
                    {
                        pic.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 0, Units.Centimeters * 3 * (k % 10 - 1), Units.Centimeters * 2, Units.Centimeters * 2);

                    }
                    if (File.Exists("./" + "20" + Student.Tables[0].Rows[k]["XH"].ToString().Substring(0,2) + "/" + Student.Tables[0].Rows[k]["XH"].ToString() + ".jpg"))
                        pic.Image = Image.FromFile("./" + "20"+Student.Tables[0].Rows[k]["XH"].ToString().Substring(0,2) + "/" + Student.Tables[0].Rows[k]["XH"].ToString()+ ".jpg");
                    data2.Objects.Add(pic);

                    BarcodeObject bar = new BarcodeObject();
                    if (k % 2 == 0)
                    {
                        bar.Bounds = new RectangleF(0, Units.Centimeters * 3 * (k % 10) + Units.Centimeters * 2.5f, Units.Centimeters * 8, Units.Centimeters * 2);
                    }
                    else
                    {
                        bar.Bounds = new RectangleF(Units.Centimeters * 9 + Units.Centimeters * 0, Units.Centimeters * 3 * (k % 10 - 1) + Units.Centimeters * 2.5f, Units.Centimeters * 8, Units.Centimeters * 2);

                    }
                    //pic.Image = Image.FromFile("100001.jpg");
                    bar.Text = Student.Tables[0].Rows[k]["ZKZH"].ToString();
                    data2.Objects.Add(bar);

                }
            }
            this.Invoke(new setStatusDelegate1(setStatus2));
            this.Invoke(new setStatusDelegate1(reportShow1));
          
        }
        private string generatesql()
        {
            string sql = "select * from [20028]";
            string sql1;
            if (comboBox1.Text == "" && comboBox2.Text == "" && comboBox3.Text == "" && comboBox4.Text == "")
            {
                //MessageBox.Show("null");
                sql = "select * from [20028]";
            }
            else 
            {
                MessageBox.Show(comboBox1.Text);
                sql = sql + "where";
                if (comboBox1.Text != "")
                    sql = sql + " substr(ZKZH,6,1)||'-'||substr(ZKZH,11,3)='" + comboBox1.Text+"'"+ " and";
                if (comboBox2.Text != "")
                {  
                    if(comboBox2.Text=="四级")
                        sql = sql + " substr(ZKZH,9,1)=" + "'1'" + " and";
                    if (comboBox2.Text == "六级")
                        sql = sql + " substr(ZKZH,9,1)=" + "'1'" + " and";
                }
                if (comboBox3.Text != "")
                {
                    if (comboBox3.Text == "日语")
                        sql = sql + " substr(ZKZH,10,1)=" + "'3'" + " and";
                    if (comboBox3.Text == "法语")
                        sql = sql + " substr(ZKZH,10,1)=" + "'7'" + " and";
                    if (comboBox3.Text == "俄语")
                        sql = sql + " substr(ZKZH,10,1)=" + "'9'" + " and";
                }
                if (comboBox4.Text != "")
                {
                    if (comboBox4.Text == "丁字沽")
                        sql = sql + " substr(ZKZH,6,1)=" + "'1'";
                    if (comboBox4.Text == "北辰")
                        sql = sql + " substr(ZKZH,6,1)=" + "'2'" ;
                    if (comboBox4.Text == "廊坊")
                        sql = sql + " substr(ZKZH,6,1)=" + "'3'";
                }
               
                
               
               
            }
            int length = sql.Length;
            if (sql.Substring(sql.Length - 3, 3).Equals("and"))
            {
                 sql1 = sql.Remove(length - 3, 3);
            }
            else
            {
                sql1 = sql;
            }
            sql1 = sql1 + "order by (substr(ZKZH,6,1)||'-'||substr(ZKZH,11,5)) ";
            MessageBox.Show(sql1);
            return sql1;
        
        }
        private void button2_Click(object sender, EventArgs e)
        {
            report.Prepare();
            report.PrintPrepared();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            report.Prepare();
            report.PrintPrepared();
        }


    }
}
