using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Npgsql;
using Dapper;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace excel_csv_read
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataTable execl_table = GetDataFromExcelByConn();
            insert_data(execl_table);

        }

        //讀CSV

        //讀excel
        System.Data.DataTable GetDataFromExcelByConn(bool hasTitle = false)
        {
            System.Windows.Forms.OpenFileDialog openFile = new System.Windows.Forms.OpenFileDialog();
            //預設路徑
            openFile.InitialDirectory = @"C:\Users\190808\Desktop";
            //不允許多個檔案
            openFile.Multiselect = false;
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.Cancel) return null;
            var filePath = openFile.FileName;
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType)) return null;

            using (DataTable ds = new DataTable())
            {

                string strCon = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", filePath);

                string strCom = " SELECT * FROM  [Sheet1$]";
                using (OleDbConnection myConn = new OleDbConnection(strCon))
                using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn))
                {
                    myConn.Open();
                    myCommand.Fill(ds);
                }
                return ds;
            }
        }

        //read csv
        void Read_csv()
        {
            System.Windows.Forms.OpenFileDialog openFile = new System.Windows.Forms.OpenFileDialog();
            //預設路徑
            openFile.InitialDirectory = @"C:\Users\190808\Desktop";
            //不允許多個檔案
            openFile.Multiselect = false; 
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.Cancel);// return null;
            var filePath = openFile.FileName;//取檔名
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType)) ;//return null
            string strconn = string.Format(@"Provider= Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Text;", filePath);

            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            StreamReader sr = new StreamReader(fs, Encoding.GetEncoding(950));  //big5的編碼為GetEncoding(950)
            string fileContent = sr.ReadToEnd();

            //將空行切割(通常的CSV是用,如遇到可將\n==>,即可)
            string[] row_str = fileContent.Split("\n".ToCharArray());

            //可在此處add class
            foreach (string acc in row_str)
            {
                string[] items = acc.Split(',');
                items[12].Substring(0, items[12].Length - 2);
            }

        }

        //新增到pgsql
        void insert_data(DataTable table)
        {
            var order_intransitstock = new List<order_intransitstock>();


            foreach (DataRow od in table.Rows)
            {
                //第一個不是空字串的才加入(後面會有很多空的也被讀到表裡面)
                if (od.ItemArray[0].ToString() != "")
                {
                    //將excel的值add到class裡面
                    order_intransitstock.Add(new order_intransitstock()
                    {
                        org = od.ItemArray[0].ToString(),
                        order_week = Convert.ToDateTime(od.ItemArray[1]),
                        customer_name = od.ItemArray[2].ToString(),
                        forecast_quantity = od.ItemArray[11].ToString() == "" ? 0 : Convert.ToInt32(od.ItemArray[11]),
                        percent_variance = od.ItemArray[11].ToString() == "" ? 0 : Convert.ToDouble(od.ItemArray[12]),
                    });
                }
            }

            //pgsql連線及insert
            using (NpgsqlConnection conn = new NpgsqlConnection(db_connectString_Ngp))
            {
                conn.Open();
                string SQL = "INSERT INTO sbdb.order_intransitstock(org, order_week, customer_name,  forecast_quantity, percent_variance)" +
                                                    "VALUES(@org,@order_week,@customer_name,@forecast_quantity,@percent_variance)";

                conn.Execute(SQL, order_intransitstock);
            }
        }

        String db_connectString_Ngp { get { return " Host=140.92.26.31;Port=5432;User ID=postgres;pwd=wwIin0901!;Database=postgres;Pooling=False; "; } }


        public class order_intransitstock
        {
            public string org { get; set; }
            public DateTime order_week { get; set; }
            public string customer_name { get; set; }
            public int forecast_quantity { get; set; }
            public double percent_variance { get; set; }
        }
    }
}
