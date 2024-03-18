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
using HtmlAgilityPack;
using System.Net;
using System.IO;
using System.Web;
using System.Collections.Specialized;
using System.Globalization;




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
            StockDataProcess STDObj = new StockDataProcess();

            STDObj.DailyKDCheck();
            ////StockIDList_string = "3305";
            //List<string> StockIDList = new List<string>(StockIDList_string.Split(';'));
            //MSSQLCls DBObj = new MSSQLCls();

            //textBox1.Text = textBox1.Text + "Start" + "：" + DateTime.Now.ToString() + Environment.NewLine;
            //Application.DoEvents();

            //string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
            //using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
            //{
            //    AutoStock_Conn.Open();

            //    StockDataProcess STDObj = new StockDataProcess();
            //    STDObj.DailyKDCheck();

            //    //string sql = "select distinct StockID from StockDailyData ";
            //    //SqlDataReader DR = DBObj.sqlDrCreater(sql, AutoStock_Conn);

            //    ////foreach (string StockID in StockIDList)
            //    //while (DR.Read())
            //    //{
            //    //    //List<string> CHK = DBObj.sqlGetValue("select top 1 StockID from StockDailyData where StockID = '" + StockID + "' ", AutoStock_Conn);

            //    //    //if (CHK[0].ToString() == "_eof")
            //    //    //{ 
            //    //        dataInput(DR["StockID"].ToString());
            //    //    //}
            //    //}
            //}


            CfunObj.showMessageBox("已完成");
    
    }


        private void StockEPSGet(string myear , string season)
        {
            //TYPEK = "pub"
            //code = ""
            //year = "112"
            //season = "01 02 03 04"
            //season = "01 02 03 04"
            //https://blog.shiangsoft.com/stock-season-report/

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument HtmlDoc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument HtmlDoc1 = new HtmlAgilityPack.HtmlDocument();
            string urlpath = "https://mops.twse.com.tw/mops/web/ajax_t163sb19?encodeURIComponent=1&step=1&firstin=1&TYPEK=sii&code=&year="+ myear + "&season="+ season;
                        
            WebClient url = new WebClient();//3分鐘
            MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));
            HtmlDoc.Load(ms, Encoding.UTF8);

            CommondFuncCls CFNOBJ = new CommondFuncCls();

            HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes("//table[@class=\"hasBorder\"]");
            MSSQLCls DBObj = new MSSQLCls();
            string connstr = "Password = what1234; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
            using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
            {
                AutoStock_Conn.Open();
                string Ins_sql = "";
                foreach (var tableNode in htnode)
                {
                    //CFNOBJ.showMessageBox("I am here");
                    var trNodes = tableNode.SelectNodes("./tr");

                    string StockID = "";
                    string StockType = "";
                    string StockEPS = "";
                    

                    foreach (var trNode in trNodes)
                    {
                        if (trNode.SelectSingleNode("td[1]") != null)
                        { 
                            StockID =  trNode.SelectSingleNode("td[1]").InnerText;
                            StockType = trNode.SelectSingleNode("td[3]").InnerText.Trim();
                            StockEPS = trNode.SelectSingleNode("td[4]").InnerText.Trim();
                            string chksql = " select StockID from  StockEPS where StockID = '" + StockID + "' and EPS_Myear = '" + myear + "' and EPS_season = '" + season + "' ";
                            List<string> chkList = DBObj.sqlGetValue(chksql, AutoStock_Conn);
                            if (chkList[0] == "_eof")
                            {
                                Ins_sql = " insert into StockEPS  (StockID,StockType,EPS_Myear,EPS_season,EPS) values ('" + StockID + "','" + StockType + "' , '" + myear + "','" + season + "','" + StockEPS + "' )";
                                DBObj.sqlExecute(Ins_sql, AutoStock_Conn);
                            }
                        }
                    }
                }
                
            }


        }

    private void StockIncomeGet(string URL_YearMonth)
    {
        MSSQLCls DBObj = new MSSQLCls();
        HttpWebRequest request = null;
        string result = null;
        int index = 0;

        string URL = "https://mops.twse.com.tw/nas/t21/sii/t21sc03_" + URL_YearMonth + ".csv";
        request = (HttpWebRequest)WebRequest.Create(URL);

        // 將 HttpWebRequest 的 Method 屬性設置為 GET
        request.Method = "GET";
        using (WebResponse response = request.GetResponse())
        {
            StreamReader sr = new StreamReader(response.GetResponseStream());
            while (!sr.EndOfStream)
            {
                result = sr.ReadLine();
                index++;

                if (index > 1)
                {
                    string[] data = result.Split(',');

                    //string Date = data[0].Replace('"', ' ').ToString();
                    string DateMonth = data[1].Replace('"', ' ').ToString().Trim();
                    string StockID = data[2].Replace('"', ' ').ToString().Trim();
                    string StockName = data[3].Replace('"', ' ').ToString().Trim();
                    string StockType = data[4].Replace('"', ' ').ToString().Trim();
                    string InCome_Month = data[5].Replace('"', ' ').ToString().Trim();
                    string InCome_LastMonth = data[6].Replace('"', ' ').ToString().Trim();
                    string InCome_LastYearMonth = data[7].Replace('"', ' ').ToString().Trim();
                    string InComeDiff_LastMonth = data[8].Replace('"', ' ').ToString().Trim();
                    string InComeDiff_InCome_LastYearMonth = data[9].Replace('"', ' ').ToString().Trim();
                    string InCome_Year = data[10].Replace('"', ' ').ToString().Trim();
                    string InCome_LastYear = data[11].Replace('"', ' ').ToString().Trim();
                    string InComeDiff_InCome_LastYear = data[12].Replace('"', ' ').ToString().Trim();

                    string connstr = "Password = what1234; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
                    using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                    {
                            AutoStock_Conn.Open();

                            MSSQLCls SQLobj = new MSSQLCls();

                            string CHECKsql = "select DateMonth from StockIncome where DateMonth = '" + DateMonth + "'  and StockID = '" + StockID + "' ";

                            List<string> eofcheck = SQLobj.sqlGetValue(CHECKsql, AutoStock_Conn);
                            string sql = "";

                            if (eofcheck[0] == "_eof")
                            {
                                sql = " insert into StockIncome (DateMonth,StockID,StockName,StockType,InCome_Month,InCome_LastMonth,InCome_LastYearMonth,InComeDiff_LastMonth,InComeDiff_InCome_LastYearMonth,InCome_Year,InCome_LastYear,InComeDiff_InCome_LastYear) " +
                                    "values ('" + DateMonth + "','" + StockID + "','" + StockName + "','" + StockType + "','" + InCome_Month + "','" + InCome_LastMonth + "','" + InCome_LastYearMonth + "','" + InComeDiff_LastMonth + "','" + InComeDiff_InCome_LastYearMonth + "','" + InCome_Year + "','" + InCome_LastYear + "','" + InComeDiff_InCome_LastYear + "') ";
                                SQLobj.sqlExecute(sql, AutoStock_Conn);
                                Thread.Sleep(100);
                            }
                            textBox1.Text = URL_YearMonth + "//" + StockID  + Environment.NewLine + textBox1.Text;
                            Application.DoEvents();
                    }
                }
                Thread.Sleep(100);
            }
        }



        }

        

        private void dataInput(string id)
        {
            CommondFuncCls CfunObj = new CommondFuncCls();
            StockDataProcess SDPObj = new StockDataProcess();
            DateTime everyMonth = new DateTime(2022, 06, 01);
            string everyMonthStr = CfunObj.DateToStringLen8(everyMonth);

            string flag_error = "go";

            while (everyMonthStr != "20231001")
            {
                everyMonthStr = CfunObj.DateToStringLen8(everyMonth);
                string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
                SqlConnection AutoStock_Conn = new SqlConnection(connstr);
                AutoStock_Conn.Open();
                MSSQLCls SQLobj = new MSSQLCls();

                //112 / 10 / 17
                string m_year = (everyMonth.Year - 1911).ToString();
                string m_month = ("0"+(everyMonth.Month).ToString()).Substring(("0" + (everyMonth.Month).ToString()).Length - 2, 2);
                string m_date = m_year + "/" + m_month;
                List<string> CHK = SQLobj.sqlGetValue("select top 1 StockID from StockDailyData where StockID = '" + id.Trim() + "' and DataDate like '" + m_date + "%' ", AutoStock_Conn);
                if (CHK[0] != "_eof") flag_error = "done";
               while (flag_error == "go")
                {
                    try
                    {
                       SDPObj.DailyDatatoDB(id, everyMonthStr);
                        //Thread.Sleep(1000);
                        Application.DoEvents();
                        flag_error = "done";
                    }
                    catch
                    {
                        textBox1.Text = id + "：" + DateTime.Now.ToString() + " SLEEP " + Environment.NewLine + textBox1.Text;
                        Application.DoEvents();
                        Thread.Sleep(185325);
                        flag_error = "go";
                    }
                }
                everyMonth = everyMonth.AddMonths(1);
                flag_error = "go";
                everyMonthStr = CfunObj.DateToStringLen8(everyMonth);
                Application.DoEvents();
                textBox1.Text = everyMonthStr + "      " + id + Environment.NewLine + textBox1.Text;
                Thread.Sleep(3000);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
