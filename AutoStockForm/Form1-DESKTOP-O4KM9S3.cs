using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AutoStockClass;
using System.Threading;
using System.Data.SqlClient;
using ClassLibrary;
using System.Xml;


namespace AutoStockForm
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommondFuncCls CfunObj = new CommondFuncCls();
            string StockIDList_string = "8131,6263,4305,3528,2414,3036,2324,3702,2886,2885,2881,1788";
            //2327,2347,2356,2357,2376,2377,2382,2385,2404,2408,2439"
            //+",2449,2451,2492,2609,2886,2915,3005,3034,3231,3532,1101,1102,1210,1216,1227,1229,1301,1303,1326,1402,1434,1476,1477,1504"
            //+ ",1536,1590,1605,1707,1717,1722,1723,1802,2002,2027,20492059,2101,2104,2105,2201,2207,2227,2231,2301,2303,2308,2313,2317,"
            //+"2324,2327,2330,2337,2344,2345,2347,2353,2354,2356,2357,2360,2371,2376,2377,2379,2382,2383,2385,2395,2404,2408,2409,2412,"
            //+"2439,2448,2449,2451,2454,2474,2492,2498,2542,2603,2606,2610,2615,2618
            //2801,2809,2812,2823,2845,2867,2881,2882,2883,"
            //"2884,2885,2886,2887,2888,2889,2890,2891,2892,2903,2912,2915,3005,3008,3023,3034,3037,3044,3045,3231,3406,3443,3481,3532,"
            // "3533,3702,3706,4904,4938,4958,5269,5522,5871,5880,6176,6213,6239,6269,6278,6409,6415,6456,6505
            //1465,6177,6292,5604,8112,1615,2527,2855,2227

            //STDObj.DailyKDCheck();


            StockDataProcess STDObj = new StockDataProcess();
            //STDObj.DailyDatatoDB("2885", "20200730");
            //STDObj.DailyBuyOvertoDB("20200730");
            //STDObj.DailyKDCheck();

            

            List<string> StockIDList = new List<string>(StockIDList_string.Split(','));
            textBox1.Text = textBox1.Text + "Start" + "：" + DateTime.Now.ToString() + Environment.NewLine;
            Application.DoEvents();
            string connstr = "Password = what1234; User ID = sa; Initial Catalog = AutoStock; Data Source = localhost; Connection Timeout = 0; Connection Lifetime = 4800; ";
            using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
            {
                MSSQLCls DBObj = new MSSQLCls();
                AutoStock_Conn.Open();
                StockDataProcess SDPObj = new StockDataProcess();
                //SDPObj.DailyBuyOvertoDB("20200714");
                //SDPObj.DailyBuyOvertoDB("20200715");
                //SDPObj.DailyBuyOvertoDB("20200716");
                //SDPObj.DailyBuyOvertoDB("20200717");
                //SDPObj.DailyBuyOvertoDB("20200720");
                //SDPObj.DailyBuyOvertoDB("20200721");

                //return;
                foreach (string Sid in StockIDList)
                {
                    List<string> isEof = DBObj.sqlGetValue("select StockID from StockDailyData where StockID = '" + Sid + "' ", AutoStock_Conn);
                    if (isEof[0] == "_eof")
                    {
                        dataInput(Sid);
                        textBox1.Text = Sid + "：" + DateTime.Now.ToString() + Environment.NewLine + textBox1.Text;
                        Application.DoEvents();
                    }
                }
            }
            textBox1.Text = textBox1.Text + "Start" + "：" + DateTime.Now.ToString() + Environment.NewLine;
            Application.DoEvents();
            //StockDataProcess STDObj = new StockDataProcess();
            //STDObj.DailyKDCheck();
            textBox1.Text = textBox1.Text + "DailyStockDataIn" + "：" + DateTime.Now.ToShortDateString() + Environment.NewLine;
            Application.DoEvents();
            STDObj.DailyStockDataIn();
            //textBox1.Text = textBox1.Text + "DailyKDCheck" + "：" + DateTime.Now.ToString() + Environment.NewLine;
            Application.DoEvents();
            //STDObj.DailyKDCheck();

            CfunObj.showMessageBox("已完成");
        }

        private Boolean DailyStockDataIn()
        {
            try
            {
                StockDataProcess SDPObj = new StockDataProcess();
                MSSQLCls SQLobj = new MSSQLCls();
                CommondFuncCls CfunObj = new CommondFuncCls();

                System.DateTime currentTime = new System.DateTime();
                currentTime = System.DateTime.Now;
                string thisday_len8 = CfunObj.DateToStringLen8(currentTime);

                string connstr = "Password = what1234; User ID = sa; Initial Catalog = AutoStock; Data Source = localhost; Connection Timeout = 0; Connection Lifetime = 4800; ";
                using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                {
                    AutoStock_Conn.Open();
                    string sqlstr = "select distinct StockID from StockDailyData";
                    DataSet dataset1 = new DataSet();
                    SqlDataReader sqlDr = SQLobj.sqlDrCreater(sqlstr, AutoStock_Conn);
                    while (sqlDr.Read())
                    {
                        SDPObj.DailyDatatoDB(sqlDr["StockID"].ToString(), thisday_len8);
                        Thread.Sleep(15000);
                        using (SqlConnection AutoStock2_Conn = new SqlConnection(connstr))
                        {
                            AutoStock2_Conn.Open();
                            string DailyKD_sp_sql = "exec DailyKD_sp '" + sqlDr["StockID"].ToString() + "',0,9";
                            SQLobj.sqlExecute(DailyKD_sp_sql, AutoStock2_Conn);
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private Boolean DailyKDCheck()
        {
            try
            {
                StockDataProcess SDPObj = new StockDataProcess();
                MSSQLCls SQLobj = new MSSQLCls();
                CommondFuncCls CfunObj = new CommondFuncCls();

                System.DateTime currentTime = new System.DateTime();
                currentTime = System.DateTime.Now;
                string thisday_len8 = CfunObj.DateToStringLen8(currentTime);

                string connstr = "Password = what1234; User ID = sa; Initial Catalog = AutoStock; Data Source = localhost; Connection Timeout = 0; Connection Lifetime = 4800; ";
                string returnMessage = "";
                using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                {
                    AutoStock_Conn.Open();
                    string sqlstr = "select distinct StockID from StockDailyData";

                    DataSet dataset1 = new DataSet();
                    SqlDataReader sqlDr = SQLobj.sqlDrCreater(sqlstr, AutoStock_Conn);
                    while (sqlDr.Read())
                    {
                        string KDsql = "select top 1 StockID,DataDate,DailyK, DailyD  FROM  StockDailyData where StockID = '" + sqlDr["StockID"].ToString() + "' "
                            + " and (DailyK >=80 or DailyD >= 80 or DailyK <= 20 or DailyD <= 20) ORDER BY DataDate DESC";
                        using (SqlConnection AutoStock2_Conn = new SqlConnection(connstr))
                        {
                            AutoStock2_Conn.Open();
                            SqlDataReader sqlDr2 = SQLobj.sqlDrCreater(KDsql, AutoStock2_Conn);
                            while (sqlDr2.Read())
                            {
                                returnMessage = returnMessage + "股票代號：" + sqlDr2["StockID"].ToString() + " KD值突破20/80  K：" + sqlDr2["DailyK"].ToString() + " D：" + sqlDr2["DailyD"].ToString() + "<br>";
                            }
                        }
                    }
                }

                if (CfunObj.isEmpty(returnMessage) == false)
                {
                    CfunObj.libHtmlMail("監控的股票KD值突破20/80", returnMessage, "cheelsu@gmail.com;downtoearth.t@gmail.com", "");
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void dataInput( string id)
        {
            CommondFuncCls CfunObj = new CommondFuncCls();
            StockDataProcess SDPObj = new StockDataProcess();
            DateTime everyMonth = new DateTime(2016, 01, 01);
            string everyMonthStr = CfunObj.DateToStringLen8(everyMonth);

            string flag_error = "go";

            while (everyMonthStr != "20200701")
            {
                everyMonthStr = CfunObj.DateToStringLen8(everyMonth);
                while (flag_error == "go")
                {
                    try
                    {
                        SDPObj.DailyDatatoDB(id, everyMonthStr);
                        Thread.Sleep(5000);
                        Application.DoEvents();
                        flag_error = "done";
                    }
                    catch
                    {
                        flag_error = "go";
                    }
                }
                everyMonth = everyMonth.AddMonths(1);
                flag_error = "go";
                everyMonthStr = CfunObj.DateToStringLen8(everyMonth);
                Application.DoEvents();
                textBox1.Text =  everyMonthStr +"      "+ id +  Environment.NewLine + textBox1.Text;
               
            }
        }
    }
}
