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


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            DataSet customerOrders = new DataSet("CustomerOrders");

            DataTable workTable = customerOrders.Tables.Add("Customers");



            DataColumn workCol = workTable.Columns.Add("CustID", typeof(Int32));
            workCol.AllowDBNull = false;
            // workCol.Unique = true;

            workTable.Columns.Add("CustLName", typeof(String));
            workTable.Columns.Add("CustFName", typeof(String));
            workTable.Columns.Add("Purchases", typeof(Double));

            workTable.Rows.Add(new Object[] { 1, "Smith", "aaa", 2.1 });
            workTable.Rows.Add(new Object[] { 2, "jack", "aaa", 2.1 });

            workTable.Rows.Add(new Object[] { 1, "Smith1", "aaa", 2.1 });
            workTable.Rows.Add(new Object[] { 2, "jack1", "aaa", 2.1 });

            workTable.Rows.Add(new Object[] { 1, "Smith11", "aaa", 2.1 });
            workTable.Rows.Add(new Object[] { 2, "jack11", "aaa", 2.1 });
            workTable.Rows.Add(new Object[] { 1, "Smith111", "aaa", 2.1 });
            workTable.Rows.Add(new Object[] { 2, "jack111", "aaa", 2.1 });



            Report report = new Report();
            report.Load("Untitled.frx");


            // register the "Products" table

            report.RegisterData(customerOrders.Tables["Customers"], "Customers");

            // enable it to use in a report

            report.GetDataSource("Customers").Enabled = true;


            GroupHeaderBand group1 = (GroupHeaderBand)report.FindObject("GroupHeader1");



            group1.Height = Units.Centimeters * 1;

            // set group condition

            group1.Condition = "[Customers.CustLName]";



            DataBand data = (DataBand)report.FindObject("data1");
            // register the "Products" table




            data.DataSource = report.GetDataSource("Customers");



            TextObject text2 = new TextObject();

            text2.Name = "Text2";

            text2.Bounds = new RectangleF(0, 0,

              Units.Centimeters * 2, Units.Centimeters * 1);

            text2.Text = "[Customers.CustLName]";

            text2.Font = new Font("Tahoma", 10, FontStyle.Bold);
            data.Objects.Add(text2);

            // report.SetParameterValue("Parameter", "[Customers.CustLName]");
            PictureObject pic = (PictureObject)report.FindObject("Picture1");
            pic.Image = Image.FromFile("100001.jpg");

            // create A4 page with all margins set to 1cm

            /*  ReportPage page1 = new ReportPage();

              page1.Name = "Page1";

              report.Pages.Add(page1);

              DataBand data1 = new DataBand();

              data1.Name = "Data1";

              data1.Height = Units.Centimeters * 0.5f;

              page1.Bands.Add(data1);

              // set data source

                         // connect databand to a group

              // group

              TextObject text2 = new TextObject();

              text2.Name = "Text2";

              text2.Bounds = new RectangleF(0, 0,

                Units.Centimeters * 2, Units.Centimeters * 1);

              text2.Text = "test";

              text2.Font = new Font("Tahoma", 10, FontStyle.Bold);
              data1.Objects.Add(text2);


              BarcodeObject bar = new BarcodeObject();
              bar.Text = "123";
              data1.Objects.Add(bar);
           





              // run the report*/
            report.Preview = this.previewControl1;
            report.Prepare();
            report.Show();

            // report.PrintPrepared();
        }
    }
}
