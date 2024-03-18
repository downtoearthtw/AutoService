using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Net.Mail;
using System.Net;
using System.IO;
using ClassLibrary;
using System.Data.SqlClient;
using System.Threading;
using System.Data;
using System.Globalization;


namespace AutoStockClass
{
    public class StockDataProcess
    {
        DataSet dataset1 = new DataSet();
        string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
        public void DailyDatatoDB(string StockID, string sdt)
        {
            DebugCls logObj = new DebugCls();
            logObj.WriteFileLog("DailyDatatoDB", "DailyDatatoDB Start", DateTime.Now.ToString(), "","");

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument HtmlDoc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument HtmlDoc1 = new HtmlAgilityPack.HtmlDocument();


            string urlpath = "";
            string urlpath_s = "";
            string urlpath_s_flag = "";
            if (StockID == "0000")
            {
                urlpath = "https://www.twse.com.tw/ZH/indicesReport/MI_5MINS_HIST?response=html&date=" + sdt;
            }
            else
            {
                urlpath = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=" + sdt + "&stockNo="+ StockID.Trim()+"&response=html";//上市
                string sdt_s = (Int16.Parse(sdt.Substring(0, 4).ToString()) - 1911).ToString() + "/" + sdt.Substring(4, 2).ToString();
                urlpath_s = "https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_print.php?l=zh-tw&d=" + sdt_s + "&stkno=" + StockID.Trim();//上櫃
            }

            logObj.WriteFileLog("DailyDatatoDB", "DailyDatatoDB urlpath", DateTime.Now.ToString(), urlpath, "");
            logObj.WriteFileLog("DailyDatatoDB", "DailyDatatoDB urlpath_s", DateTime.Now.ToString(), urlpath_s, "");

            // WebClient url = new WebClientTo(1800000);//3分鐘
            WebClient url = new WebClient();//3分鐘
            MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));

            HtmlDoc.Load(ms, Encoding.UTF8);
            //HtmlDoc1.LoadHtml(HtmlDoc.DocumentNode.SelectSingleNode("/html/body/div/table/tbody/tr[11]").InnerHtml);

            MSSQLCls SQLobj = new MSSQLCls();
            CommondFuncCls CfunObj = new CommondFuncCls();

            HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/div/table/tbody/tr");

            if (CfunObj.isEmpty(htnode))
            {
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                ms = new MemoryStream(url.DownloadData(urlpath_s));

                HtmlDoc.Load(ms, Encoding.UTF8);
                htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/table/tbody/tr");
                urlpath_s_flag = "1";
            }

            if (CfunObj.isEmpty(htnode))
            {
                logObj.WriteFileLog("DailyDatatoDB", "DailyDatatoDB ", DateTime.Now.ToString(), "上市櫃都取不到資料", "");
                return;
            }


           string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
            using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
            {
                AutoStock_Conn.Open();
                string insSql = "";
                string DelSql = "";
                string TD_Date = "";

                foreach (HtmlNode currTR in htnode)
                {
                    string TD_PriceH = "";
                    string TD_PriceL = "";
                    string TD_PriceD = "";
                    string TD_PriceS = "";
                    string TD_Deal = "0";

                    if (StockID == "0000")
                    {
                        TD_Date = currTR.SelectSingleNode("td[1]").InnerText;
                        TD_PriceH = currTR.SelectSingleNode("td[3]").InnerText.Replace(",", "");
                        TD_PriceL = currTR.SelectSingleNode("td[4]").InnerText.Replace(",", ""); ;
                        TD_PriceD = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", ""); ;
                    }
                    else
                    {
                        if (urlpath_s_flag == "")//上市
                        {
                            TD_Date = currTR.SelectSingleNode("td[1]").InnerText;
                            TD_PriceH = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", "");
                            TD_PriceL = currTR.SelectSingleNode("td[6]").InnerText.Replace(",", "");
                            TD_PriceD = currTR.SelectSingleNode("td[7]").InnerText.Replace(",", "");
                            TD_PriceS = currTR.SelectSingleNode("td[4]").InnerText.Replace(",", ""); ;
                            TD_Deal = currTR.SelectSingleNode("td[2]").InnerText.Replace(",", "");
                            TD_Deal = Math.Round((float.Parse(TD_Deal) / 1000)).ToString();
                        }
                        else//上櫃
                        {
                            TD_Date = currTR.SelectSingleNode("td[1]").InnerText;
                            TD_PriceH = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", "");
                            TD_PriceL = currTR.SelectSingleNode("td[6]").InnerText.Replace(",", "");
                            TD_PriceD = currTR.SelectSingleNode("td[7]").InnerText.Replace(",", "");
                            TD_PriceS = currTR.SelectSingleNode("td[4]").InnerText.Replace(",", ""); ;
                            TD_Deal = currTR.SelectSingleNode("td[2]").InnerText.Replace(",", "");
                        }
                    }
                    if (CfunObj.IsNumeric(TD_PriceH) && CfunObj.IsNumeric(TD_PriceL) && CfunObj.IsNumeric(TD_PriceD))
                    {
                        insSql = insSql + "  insert into StockDailyData (StockID,DataDate,PriceH,PriceL,PriceD,PriceS,DealNum) values ('" + StockID + "','" + TD_Date + "'," + TD_PriceH + "," + TD_PriceL + "," + TD_PriceD + "," + TD_PriceS + "," + TD_Deal + ")";
                        //insSql = insSql + " insert into StockParam (StockID,DataDate,CV) values ('"+ StockID + "','" + TD_Date + "',[dbo].[GetCV]('" + StockID + "'))";
                    }
                }

                if (CfunObj.isEmpty(insSql) == false)
                {
                    string TD_Date_M = TD_Date.Substring(0, 6);
                    DelSql = "delete StockDailyData where StockID = '" + StockID + "' and  DataDate like '" + TD_Date_M + "%' ";
                    SQLobj.sqlExecute(DelSql, AutoStock_Conn);
                    SQLobj.sqlExecute(insSql, AutoStock_Conn);
                }
            }
        }


        public Boolean DatatoDB_singleDay(string sdt)
        {
            DebugCls logObj = new DebugCls();
            logObj.WriteFileLog("DatatoDB_singleDay", "DatatoDB_singleDay Start", DateTime.Now.ToString(), "", "");
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                HtmlAgilityPack.HtmlDocument HtmlDoc = new HtmlAgilityPack.HtmlDocument();

                HtmlAgilityPack.HtmlDocument HtmlDoc1 = new HtmlAgilityPack.HtmlDocument();


                string urlpath = "";
                string urlpath_s = "";
                string urlpath_s_flag = "";

                //urlpath = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=" + sdt + "&stockNo=" + StockID.Trim() + "&response=html";//上市
                urlpath = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date=" + sdt + "&type=ALLBUT0999&response=html";

                string sdt_s = (Int16.Parse(sdt.Substring(0, 4).ToString()) - 1911).ToString() + "/" + sdt.Substring(4, 2).ToString();
                string sdt_s_i = (Int16.Parse(sdt.Substring(0, 4).ToString()) - 1911).ToString() + "/" + sdt.Substring(4, 2).ToString() + "/" + sdt.Substring(6, 2).ToString(); ;
                //urlpath_s = "https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_print.php?l=zh-tw&d=" + sdt_s + "&stkno=" + StockID.Trim();//上櫃
                urlpath_s = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&o=htm&d=" + sdt_s_i + "&se=EW&s=0,asc,0";

                logObj.WriteFileLog("DatatoDB_singleDay", "DatatoDB_singleDay urlpath", DateTime.Now.ToString(), urlpath, "");
                logObj.WriteFileLog("DatatoDB_singleDay", "DatatoDB_singleDay urlpath_s", DateTime.Now.ToString(), urlpath_s, "");

                 //WebClient url = new WebClientTo(180000);//3分鐘
                WebClient url = new WebClient();
                byte[] dbytes = url.DownloadData(urlpath);
                string responseStr = System.Text.Encoding.UTF8.GetString(dbytes);

                byte[] dbyte_S = url.DownloadData(urlpath_s);
                string responseStr_S = System.Text.Encoding.UTF8.GetString(dbyte_S);

                HtmlDoc.LoadHtml(responseStr);

                MSSQLCls SQLobj = new MSSQLCls();
                CommondFuncCls CfunObj = new CommondFuncCls();

                string TD_ID = "";
                string TD_NAME = "";
                string TD_DealQTY = "";//成交股數
                string TD_DealMoney = "";//成交金額
                string TD_PriceS = "";//開盤價
                string TD_PriceH = "";//最高價
                string TD_PriceL = "";//最低價
                string TD_PriceD = "";//收盤價
                string insSql = "";

                for (int i = 1; i < 2000; i++) //先取上市
                { 
                    HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/div/table[9]/tbody/tr[" + i.ToString() + "]");
                    if ( CfunObj.isEmpty(htnode)) break;

                    foreach (HtmlNode currTR in htnode)
                    {
                        TD_ID = "";
                        TD_NAME = "";
                        TD_DealQTY = "";//成交股數
                        TD_DealMoney = "";//成交金額
                        TD_PriceS = "";//開盤價
                        TD_PriceH = "";//最高價
                        TD_PriceL = "";//最低價
                        TD_PriceD = "";//收盤價
                    
                        TD_ID = currTR.SelectSingleNode("td[1]").InnerText;
                        TD_NAME = currTR.SelectSingleNode("td[2]").InnerText;
                        TD_DealQTY = currTR.SelectSingleNode("td[3]").InnerText.Replace(",", ""); ;
                        TD_DealMoney = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", ""); ;
                        TD_PriceS = currTR.SelectSingleNode("td[6]").InnerText.Replace(",", ""); ;
                        TD_PriceH = currTR.SelectSingleNode("td[7]").InnerText.Replace(",", "");
                        TD_PriceL = currTR.SelectSingleNode("td[8]").InnerText.Replace(",", "");
                        TD_PriceD = currTR.SelectSingleNode("td[9]").InnerText.Replace(",", "");

                        TD_DealQTY = Math.Round((float.Parse(TD_DealQTY) / 1000)).ToString();
                        if (CfunObj.IsNumeric(TD_PriceH) && CfunObj.IsNumeric(TD_PriceL) && CfunObj.IsNumeric(TD_PriceD) && TD_ID.Length==4)
                        {
                            insSql = insSql + "  insert into StockDailyData (StockID,StockName,DataDate,PriceH,PriceL,PriceD,PriceS,DealNum) " +
                                "values ('" + TD_ID + "','" + TD_NAME + "','" + sdt_s_i + "'," + TD_PriceH + "," + TD_PriceL + "," + TD_PriceD + "," + TD_PriceS + "," + TD_DealQTY + ")";
                        }
                    }
                }
                    

                HtmlDoc.LoadHtml(responseStr_S);

                for (int i = 1; i < 2000; i++) //取上櫃
                {
                    HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/table/tbody/tr[" + i.ToString() + "]");
                    if (CfunObj.isEmpty(htnode)) break;

                    TD_ID = "";
                    TD_NAME = "";
                    TD_DealQTY = "";//成交股數
                    TD_DealMoney = "";//成交金額
                    TD_PriceS = "";//開盤價
                    TD_PriceH = "";//最高價
                    TD_PriceL = "";//最低價
                    TD_PriceD = "";//收盤價
                
                    foreach (HtmlNode currTR in htnode)
                    {
                        TD_ID = currTR.SelectSingleNode("td[1]").InnerText; 
                        TD_NAME = currTR.SelectSingleNode("td[2]").InnerText; 
                        TD_PriceD = currTR.SelectSingleNode("td[3]").InnerText.Replace(",", ""); ;
                        TD_PriceS = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", ""); ;
                        TD_PriceH = currTR.SelectSingleNode("td[6]").InnerText.Replace(",", "");
                        TD_PriceL = currTR.SelectSingleNode("td[7]").InnerText.Replace(",", "");
                        TD_DealQTY = currTR.SelectSingleNode("td[8]").InnerText.Replace(",", "");
                    }
                    if (CfunObj.IsNumeric(TD_PriceH) && CfunObj.IsNumeric(TD_PriceL) && CfunObj.IsNumeric(TD_PriceD) && TD_ID.Length == 4)
                    {
                        insSql = insSql + "  insert into StockDailyData (StockID,StockName,DataDate,PriceH,PriceL,PriceD,PriceS,DealNum) " +
                            "values ('" + TD_ID + "','" + TD_NAME + "','" + sdt_s_i + "'," + TD_PriceH + "," + TD_PriceL + "," + TD_PriceD + "," + TD_PriceS + "," + TD_DealQTY + ")";
                    }
                }
                if (CfunObj.isEmpty(insSql) == false)
                {
                    using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                    {
                        AutoStock_Conn.Open();
                        SQLobj.sqlExecute("delete StockDailyData where DataDate = '"+ sdt_s_i + "' ", AutoStock_Conn);
                        SQLobj.sqlExecute(insSql, AutoStock_Conn);
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                logObj.ErrHandle("DatatoDB_singleDay", e);
                return false;
            }
        }


        public void DailyBuyOvertoDB(string sdt)
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument HtmlDoc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument HtmlDoc1 = new HtmlAgilityPack.HtmlDocument();


            string urlpath = "";
            string urlpath_s = "";
            string urlpath_s_flag = "";
            urlpath = "https://www.twse.com.tw/fund/T86?response=html&date=" + sdt + "&selectType=ALLBUT0999";


            //WebClient url = new WebClientTo(1800000);//3分鐘
            WebClient url = new WebClient();//3分鐘
            MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));

            HtmlDoc.Load(ms, Encoding.UTF8);
            //HtmlDoc1.LoadHtml(HtmlDoc.DocumentNode.SelectSingleNode("/html/body/div/table/tbody/tr[11]").InnerHtml);

            MSSQLCls SQLobj = new MSSQLCls();
            CommondFuncCls CfunObj = new CommondFuncCls();

            HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/div/table/tbody/tr");



            if (CfunObj.isEmpty(htnode))
            {
                return;
            }

            string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
            using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
            {
                AutoStock_Conn.Open();
                string insSql = "";
                string DelSql = "";
                string TD_Date = "";

                foreach (HtmlNode currTR in htnode)
                {
                    string TD_StockID = "";
                    string TD_ChinaBuyOver = "";
                    string TD_ForeignBuyOver = "";
                    string TD_ITBuyOver = "";
                    string TD_SelfBuyOver = "";
                    string TD_VIXBuyOver = "";

                    TD_StockID = currTR.SelectSingleNode("td[1]").InnerText;
                    TD_ChinaBuyOver = currTR.SelectSingleNode("td[5]").InnerText.Replace(",", "");//外陸資
                    TD_ForeignBuyOver = currTR.SelectSingleNode("td[6]").InnerText.Replace(",", ""); //外資 自營  
                    TD_ITBuyOver = currTR.SelectSingleNode("td[8]").InnerText.Replace(",", "");//投信
                    TD_SelfBuyOver = currTR.SelectSingleNode("td[12]").InnerText.Replace(",", "");//自營
                    TD_VIXBuyOver = currTR.SelectSingleNode("td[15]").InnerText.Replace(",", "");//避險


                    if (CfunObj.IsNumeric(TD_ChinaBuyOver) && CfunObj.IsNumeric(TD_ForeignBuyOver) && CfunObj.IsNumeric(TD_ITBuyOver) && CfunObj.IsNumeric(TD_SelfBuyOver) && CfunObj.IsNumeric(TD_VIXBuyOver))
                    {
                        insSql = insSql + "  insert into StockBuyOver (StockID,DataDate,ChinaBuyOver,ForeignBuyOver,ITBuyOver,SelfBuyOver,VIXBuyOver) "
                            + "values ('" + TD_StockID + "','" + sdt + "'," + TD_ChinaBuyOver + "," + TD_ForeignBuyOver + "," + TD_ITBuyOver + "," + TD_SelfBuyOver + "," + TD_VIXBuyOver + ")";
                    }
                }

                if (CfunObj.isEmpty(insSql) == false)
                {
                    //string TD_Date_M = TD_Date.Substring(0, 6);
                    DelSql = "delete StockBuyOver where DataDate like '" + sdt + "%' ";
                    SQLobj.sqlExecute(DelSql, AutoStock_Conn);
                    SQLobj.sqlExecute(insSql, AutoStock_Conn);
                }
            }

        }

        public Boolean DailyStockDataIn()
        {
            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("DailyStockDataIn", "DailyStockDataIn Start", DateTime.Now.ToString(), "", "");
            try
            {
                StockDataProcess SDPObj = new StockDataProcess();
                MSSQLCls SQLobj = new MSSQLCls();
                CommondFuncCls CfunObj = new CommondFuncCls();

                System.DateTime currentTime = new System.DateTime();
                currentTime = System.DateTime.Now.AddDays(-1); ;
                string thisday_len8 = CfunObj.DateToStringLen8(currentTime);

                SDPObj.DailyBuyOvertoDB(thisday_len8);//取證交所法人買賣超

                string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
                using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                {
                    AutoStock_Conn.Open();
                    DateTime today = DateTime.Now.AddDays(-1);
                    int nowHourInt = today.Hour;

                    CultureInfo culture = new CultureInfo("zh-TW");
                    culture.DateTimeFormat.Calendar = new TaiwanCalendar();
                    string todayMG = today.ToString("yyy/MM/dd", culture);
                    string sqlstr = "";
                    

                    LogObj.WriteFileLog("DailyStockDataIn", "DatatoDB_singleDay todayMG", DateTime.Now.ToString(), thisday_len8, "");
                    if (today.DayOfWeek.ToString() != "Sunday" && today.DayOfWeek.ToString() != "Saturday") 
                    { 
                        SDPObj.DatatoDB_singleDay(thisday_len8);//取當天交易資料
                    }


                    sqlstr = "select distinct StockID , max(DataDate) maxDataDate from StockDailyData group by StockID";
                    SqlDataReader sqlDr = SQLobj.sqlDrCreater(sqlstr, AutoStock_Conn);
                    string flag_error = "go";
                    while (sqlDr.Read())
                    {
                        while (flag_error == "go")
                        {
                            try
                            {
                                if (today.DayOfWeek.ToString() == "Sunday")//週日取全月，以免有漏的
                                {
                                    LogObj.WriteFileLog("DailyStockDataIn", "DailyStockDataIn ", DateTime.Now.ToString(), "取證交所每天各股成交資料", "");
                                    SDPObj.DailyDatatoDB(sqlDr["StockID"].ToString(), thisday_len8);//取證交所每天各股成交資料
                                    Thread.Sleep(5000);
                                }

                                using (SqlConnection AutoStock2_Conn = new SqlConnection(connstr))
                                {
                                    AutoStock2_Conn.Open();
                                    string DailyKD_sp_sql = "exec DailyKD_sp '" + sqlDr["StockID"].ToString() + "',0,9";
                                    SQLobj.sqlExecute(DailyKD_sp_sql, AutoStock2_Conn);
                                }
                                flag_error = "done";
                            }
                            catch
                            {
                                flag_error = "go";
                            }
                        }
                        flag_error = "go";
                    }
                }

                LogObj.WriteFileLog("DailyStockDataIn", "DailyStockDataIn Done", DateTime.Now.ToString(), "", "");
                return true;
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("DailyStockDataIn", e);
                return false;
            }
        }

        public Boolean DailyKDCheck()
        {
            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("DailyKDCheck", "DailyKDCheck Start", DateTime.Now.ToString(), "", "");
            string StockID = "";

            StockDataProcess SDPObj = new StockDataProcess();
            MSSQLCls SQLobj = new MSSQLCls();
            CommondFuncCls CfunObj = new CommondFuncCls();

            System.DateTime currentTime = new System.DateTime();
            currentTime = System.DateTime.Now;
            string thisday_len8 = CfunObj.DateToStringLen8(currentTime);

            
            string returnMessage = "";
            string KDOnlyMessage = "<P>";
            string tenP_message = "<P>";
            string UP25_message = "<P>";
            string Down25_message = "<P>";
            string DaysAvg_message = "<P>";
            string DDUp_message = "<P>";

            DateTime everyMonth = DateTime.Now.AddMonths(-1);
            string mYear = (everyMonth.Year - 1911).ToString();
            string mMonth = (everyMonth.Month).ToString();
            string URL_YearMonth = mYear + "/" + mMonth;

            string season = "";
            if (mMonth == "1" || mMonth == "2" || mMonth == "3")
            {
                season = "04";
            }

            if (mMonth == "4" || mMonth == "5" || mMonth == "6")
            {
                season = "01";
            }

            if (mMonth == "7" || mMonth == "8" || mMonth == "9")
            {
                season = "02";
            }

            if (mMonth == "10" || mMonth == "11" || mMonth == "12")
            {
                season = "03";
            }

            try
            {
                using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                {
                    AutoStock_Conn.Open();
                    //string sqlstr = "select A.*,CVS from ( " +
                    //    " select distinct StockID from( " +
                    //    "        select* from (select ROW_NUMBER() over(partition by StockID order by DataDate desc) AS Row, *from StockDailyData   ) SS where Row < 30 " +
                    //    "	) sss    group by StockID having avg(DealNum) > 200000  " +
                    //    " ) A inner join (    select StockID, CVS from(select ROW_NUMBER() over(partition by StockID order by DataDate desc) AS Row,[dbo].[GetCV](StockID) CVS, * " +
                    //    " from StockDailyData  ) SS where Row = 1    and CVS < 3) B on A.StockID = B.StockID ";

                    string sqlstr = " exec GetEveryDayData ";

                    string Flag0000 = "";
                    
                    int countDDUp = 0;
                    string Days13AvgString = "";
                    string Days8AvgString = "";
                    string Days3AvgString = "";
                    string HeadString = "";
                    string StockName = "";

                    SqlConnection GetValue_Conn = new SqlConnection(connstr);
                    GetValue_Conn.Open();

                    DataSet dataset1 = new DataSet();
                    SqlDataReader sqlDr = SQLobj.sqlDrCreater(sqlstr, AutoStock_Conn);
                    while (sqlDr.Read())
                    {
                        StockID = sqlDr["StockID"].ToString();
                        StockName = sqlDr["StockName"].ToString();
                        LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "", "");

                        string Trending = "";
                        using (SqlConnection ShowTrend = new SqlConnection(connstr))
                        {
                            ShowTrend.Open();
                            string chk_sql = "select [dbo].[ShowTrend]('" + StockID.Trim() + "')";
                            List<string> LiChk = SQLobj.sqlGetValue(chk_sql, ShowTrend);
                            Trending = LiChk[0];
                        }


                        #region "計算KD"
                        /*
                        string KDsql = " select top 1 SData.serialInt,SData.StockID,SData.DataDate,SData.DailyK, SData.DailyD, SData.PriceD , "
                            + " isnull(Dividend,0)*40 DIV40 , isnull(Dividend,0)*20 DIV20, isnull(Dividend,0) * 10 DIV10 , isnull(SData.DealNum,0) DealNum "
                            + " FROM  StockDailyData SData left join StockDividend SDiv on SData.StockID = SDiv.StockID "
                            + " where SData.StockID ='" + sqlDr["StockID"].ToString() + "' ORDER BY SData.StockID , SData.DataDate DESC ";

                        string DealNumsql = " select top 1 SData.serialInt,SData.StockID,SData.DataDate , isnull(SData.DealNum,0) DealNum "
                            + " FROM  StockDailyData SData left join StockDividend SDiv on SData.StockID = SDiv.StockID "
                            + " where SData.StockID ='" + sqlDr["StockID"].ToString() + "' ORDER BY SData.StockID , SData.DataDate DESC ";

                        List<string> ListDeaNum = SQLobj.sqlGetValue(DealNumsql, GetValue_Conn);


                        float DealNum = 0;
                        if (ListDeaNum[0] != "_eof")
                        {
                            if (float.TryParse(ListDeaNum[2].ToString(), out DealNum) == false)
                            {
                                DealNum = 0;
                            }
                        }

                        string MA60String = " select PD, Avg(PriceD)MA60 from( "
                            + " select Top 60 AA.PriceD PD, BB.* "
                            + " from StockDailyData BB "
                            + " inner join ( "
                            + "     select Top 1 * from StockDailyData "
                            + "     where StockID = '" + sqlDr["StockID"].ToString() + "'  "
                            + "     order by DataDate desc "
                            + " ) AA on AA.serialInt >= BB.serialInt "
                            + " where BB.StockID = '" + sqlDr["StockID"].ToString() + "'  "
                            + " order by BB.DataDate desc "
                            + " ) TB60 "
                            + " group by PD ";

                        List<string> ListMA60 = SQLobj.sqlGetValue(MA60String, GetValue_Conn);

                        float PD = 0;
                        float MA60 = 0;

                        if (ListMA60[0] != "_eof")
                        {
                            PD = float.Parse(ListMA60[0].ToString());
                            MA60 = float.Parse(ListMA60[1].ToString());
                        }

                        
                        using (SqlConnection AutoStock2_Conn = new SqlConnection(connstr))
                        {
                            AutoStock2_Conn.Open();
                            SqlDataReader sqlDr2 = SQLobj.sqlDrCreater(KDsql, AutoStock2_Conn);
                            while (sqlDr2.Read())
                            {
                                using (SqlConnection AutoStock3_Conn = new SqlConnection(connstr))
                                {
                                    AutoStock3_Conn.Open();


                                    //K或D>20 
                                    if ((float.Parse(sqlDr2["DailyK"].ToString()) <= 20 || float.Parse(sqlDr2["DailyD"].ToString()) <= 20))
                                    {
                                        //然後 K > D 時，判斷上一筆資料是不是 D > K，如果是就是 黃金交叉
                                        if (float.Parse(sqlDr2["DailyD"].ToString()) < float.Parse(sqlDr2["DailyK"].ToString()))
                                        {
                                            string KDsql8020 = "select top 1 DailyD - DailyK FROM  StockDailyData "
                                            + "where StockID = '" + sqlDr["StockID"].ToString() + "' "
                                            + " and serialInt < '" + sqlDr2["serialInt"].ToString() + "' ORDER BY DataDate DESC";

                                            List<string> CKList = SQLobj.sqlGetValue(KDsql8020, AutoStock3_Conn);
                                            if (float.Parse(CKList[0]) > 0)
                                            {
                                                returnMessage = returnMessage + "日期：" + sqlDr2["DataDate"].ToString() + "，股票代號：" + sqlDr2["StockID"].ToString() + "形成黃金交叉，K：" + sqlDr2["DailyK"].ToString() + "，D：" + sqlDr2["DailyD"].ToString() + "<BR>"
                                                + "，收盤價：" + PD.ToString() + "，MA60：" + MA60.ToString()
                                                + "，股利比：" + sqlDr2["DIV40"].ToString() + "/" + sqlDr2["DIV20"].ToString() + "/" + sqlDr2["DIV10"].ToString() + "<br>";

                                                KDOnlyMessage = KDOnlyMessage + "日期：" + sqlDr2["DataDate"].ToString() + "，股票代號：" + sqlDr2["StockID"].ToString() + "，形成黃金交叉，K：" + sqlDr2["DailyK"].ToString() + "，D：" + sqlDr2["DailyD"].ToString() + "<BR>"
                                                    + "，收盤價：" + PD.ToString() + "，MA60：" + MA60.ToString()
                                                    + "，股利比：" + sqlDr2["DIV40"].ToString() + "/" + sqlDr2["DIV20"].ToString() + "/" + sqlDr2["DIV10"].ToString() + "<br>";
                                            }
                                        }
                                        else
                                        {
                                            if (float.Parse(sqlDr2["DailyD"].ToString()) - float.Parse(sqlDr2["DailyK"].ToString()) <= 7 && float.Parse(sqlDr2["DailyD"].ToString()) - float.Parse(sqlDr2["DailyK"].ToString()) > 0)
                                            {
                                                if (PD >= MA60)
                                                {
                                                    //KDOnlyMessage = KDOnlyMessage + "日期：" + sqlDr2["DataDate"].ToString() + "，股票代號：" + sqlDr2["StockID"].ToString() + "，KD值小於20 "
                                                    //+ "，收盤價：" + PD.ToString() + "，MA60：" + MA60.ToString() 
                                                    //+ "，股利比：" + sqlDr2["DIV40"].ToString() + "/" + sqlDr2["DIV20"].ToString() + "/" + sqlDr2["DIV10"].ToString() + "<br>";
                                                }
                                                //returnMessage = returnMessage + "日期：" + sqlDr2["DataDate"].ToString() + "，股票代號：" + sqlDr2["StockID"].ToString() + "，KD值小於20，K：" + sqlDr2["DailyK"].ToString() + "，D：" + sqlDr2["DailyD"].ToString() + "<br>";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        */
                        #endregion

                        string price_vol_flag = "";

                        #region "計算MACD"
                        LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "MACD", "");
                        using (SqlConnection MACD_Conn = new SqlConnection(connstr))
                        {
                            MACD_Conn.Open();
                            string MA2_sql = "exec MACD '" + StockID + "' ";
                            SQLobj.sqlExecute(MA2_sql, MACD_Conn);

                            string MACD_string = "select ROW_NUMBER() over(order by DataDate desc) AS RowI , * from ( " +
                                " select top 4 isnull(DIF,0) DIF,isnull(DEA,0) DEA,isnull(EMA_Q,0) EMA_Q,isnull(EMA_L,0) EMA_L,DataDate,StockID,StockName" +
                                " from dbo.StockDailyData where StockID = '" + StockID + "' order by DataDate desc ) macd_dr ";

                            SqlDataReader MACD_Dr = SQLobj.sqlDrCreater(MACD_string, MACD_Conn);
                            float DIF = 0;
                            float DEA = 0;

                            float Last_DIF = 0;
                            float Last_DEA = 0;

                            float Last_DIF_1 = 0;
                            float Last_DEA_1 = 0;

                            float Last_DIF_2 = 0;
                            float Last_DEA_2 = 0;
                            string MACDstring = "";
                            

                            while (MACD_Dr.Read())
                            {
                                LogObj.WriteFileLog("MACD", MACDstring, DateTime.Now.ToString(), "StockID", StockID);

                                if (MACD_Dr["RowI"].ToString() == "1")
                                {
                                    LogObj.WriteFileLog("MACD", "A", MACD_Dr["DIF"].ToString(), MACD_Dr["DEA"].ToString(), StockID);
                                    DIF = float.Parse(MACD_Dr["DIF"].ToString());
                                    DEA = float.Parse(MACD_Dr["DEA"].ToString());
                                }

                                if (MACD_Dr["RowI"].ToString() == "2")
                                {
                                    LogObj.WriteFileLog("MACD", "B", DateTime.Now.ToString(), "StockID", StockID);
                                    Last_DIF = float.Parse(MACD_Dr["DIF"].ToString());
                                    Last_DEA = float.Parse(MACD_Dr["DEA"].ToString());
                                }

                                if (MACD_Dr["RowI"].ToString() == "3")
                                {
                                    LogObj.WriteFileLog("MACD", "B", DateTime.Now.ToString(), "StockID", StockID);
                                    Last_DIF_1 = float.Parse(MACD_Dr["DIF"].ToString());
                                    Last_DEA_1 = float.Parse(MACD_Dr["DEA"].ToString());
                                }

                                if (MACD_Dr["RowI"].ToString() == "4")
                                {
                                    LogObj.WriteFileLog("MACD", "B", DateTime.Now.ToString(), "StockID", StockID);
                                    Last_DIF_2 = float.Parse(MACD_Dr["DIF"].ToString());
                                    Last_DEA_2 = float.Parse(MACD_Dr["DEA"].ToString());
                                    LogObj.WriteFileLog("MACD", "B-1", DateTime.Now.ToString(), "StockID", StockID);
                                }

                            }

                            if (Trending == "PASS") //VEGAS趨勢向上
                            { 
                                if (DIF < 0 && DEA < 0 )
                                {
                                    if (DIF > DEA && Last_DIF <= Last_DEA )
                                    {
                                        MACDstring = MACDstring + "  黃金交叉";
                                        price_vol_flag = "Y";
                                    }
                                }
                            }

                            if (MACDstring != "")
                            {
                                returnMessage = returnMessage +  StockID + "    " + StockName  + "    MACD" + MACDstring + StockIncomeMessage(StockID, URL_YearMonth, mYear, season)  + "<BR>";
                                price_vol_flag = "Y";
                            }
                        }
                        #endregion


                        #region "VCP型態辨識"
                        ////LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "VCP", "");
                        ////if (Trending == "PASS") //VEGAS趨勢向上
                        ////{
                        //    using (SqlConnection VCP_Conn = new SqlConnection(connstr))
                        //    {
                        //        VCP_Conn.Open();
                        //        string VCP_sql = "exec dbo.VCP  '" + StockID.Trim() + "' ";
                        //        List<string> VCPCheck = SQLobj.sqlGetValue(VCP_sql, VCP_Conn);
                        //        if (VCPCheck[0] != "_eof")
                        //        {
                        //            if (VCPCheck[2] != "")
                        //            {
                        //                returnMessage = returnMessage + StockID + VCPCheck[0].ToString() + "  " + VCPCheck[1] + "   頸線：" + VCPCheck[2] + StockIncomeMessage(StockID, URL_YearMonth, mYear, season) + " <BR>";
                        //                price_vol_flag = "Y";
                        //            }
                        //        }
                        //    }
                        ////}
                        #endregion


                        #region "均線穿越"
                        LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "均線穿越", "");
                        using (SqlConnection MA2_Conn = new SqlConnection(connstr))
                        {
                            MA2_Conn.Open();
                            string MA2_sql = "exec get_2_MA '" + StockID + "',5,21 ";

                            List<string> MA2_CHECK = SQLobj.sqlGetValue(MA2_sql, MA2_Conn);
                            if (MA2_CHECK[0] == "PASS")
                            {
                                returnMessage = returnMessage + StockID.ToString() + "    " + StockName + " 三天內5日線穿越布靈中線及兩條均線" + StockIncomeMessage(StockID, URL_YearMonth,mYear,season)  + "<BR>";
                                price_vol_flag = "Y";
                            }
                        }
                        #endregion


                        #region "連三天量增價漲"
                        LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "量價", "");
                        if (Trending == "PASS") //VEGAS趨勢向上
                        {
                            using (SqlConnection DailyPriceDiff_Conn = new SqlConnection(connstr))
                            {
                                DailyPriceDiff_Conn.Open();
                                string DailyPriceDiff_sql = "exec DailyPriceDiff_sp '" + StockID + "' ";
                                List<string> ListPriceDiff = SQLobj.sqlGetValue(DailyPriceDiff_sql, DailyPriceDiff_Conn);
                                if (ListPriceDiff[0] != "_eof")
                                {
                                    returnMessage = returnMessage + StockID + "    " + StockName.ToString() + ListPriceDiff[0] + StockIncomeMessage(StockID, URL_YearMonth, mYear, season) + "<BR>";
                                    price_vol_flag = "Y";
                                }
                            }
                        }
                        #endregion

                        #region "檢查均線糾纏"
                        if (price_vol_flag == "Y" )
                        {
                            LogObj.WriteFileLog("DailyKDCheck", StockID, DateTime.Now.ToString(), "均線糾纏", "");

                            using (SqlConnection CV_Conn = new SqlConnection(connstr))
                            {
                                CV_Conn.Open();
                                string CVstring = "select [dbo].[GetCV]('" + StockID.Trim() + "') ";
                                List<string> ListCV = SQLobj.sqlGetValue(CVstring, CV_Conn);
                                if (ListCV[0] != "_eof" && float.Parse(ListCV[0]) < 1.5 )
                                {
                                    returnMessage = returnMessage + StockID + "    " + StockName.ToString() + "    均線糾纏 CV：" + ListCV[0].ToString() + "<BR>";
                                }
                            }
                        }
                        #endregion

                    }

                    //Down25_message = "5日線在25日線之下" + Down25_message;
                    //returnMessage = returnMessage + UP25_message + "<BR>";
                    //returnMessage = returnMessage + Down25_message + "<BR>";

                    if (CfunObj.isEmpty(returnMessage) == false)
                    {
                        returnMessage = "<FONT COLOR='RED'><B>停利指標：1.DMI趨勢未能創新高，同時OBV也未能創新高</B></FONT><BR>" +
                            "<FONT COLOR='RED'><B>　　　　　2.同樣的高點，摸第二次破不了</B></FONT><BR>" +
                            "<FONT COLOR='RED'><B>　　　　　3.觸發以上，加上OBV走弱</B></FONT><BR>" +
                            "<FONT COLOR='RED'><B>　　　　　4.或是爆量不漲，長上影線或留十字線</B></FONT><BR>" +
                            "<FONT COLOR='GREEN'><B>停損指標：破前低到隔天中午1點回不來，或是回來了又摸到前低</B></FONT><BR>" +
                            "<FONT COLOR='BLUE'><B>注意：布林上緣禁止追高加碼</B></FONT><BR>" +
                            "<FONT COLOR='BLUE'><B>　　　OBV急上急下的不要碰</B></FONT><BR>" +
                            "================================================<BR>" +
                            returnMessage;
                        LogObj.WriteFileLog("returnMessage", returnMessage, DateTime.Now.ToString(), "", "");
                        CfunObj.libHtmlMail("符合指標的標的", returnMessage, "cheelsu@gmail.com;downtoearth.tw@gmail.com;greencwi@gmail.com;hunter_chang@titanlight.com", "");
                    }

                    LogObj.WriteFileLog("DailyKDCheck", "DailyKDCheck Done", DateTime.Now.ToString(), "", "");
                    return true;
                }
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("DailyKDCheck", e);
                return false;
            }
        }

        public Boolean StockIncomeGet(string URL_YearMonth)
        {
            MSSQLCls DBObj = new MSSQLCls();
            HttpWebRequest request = null;

            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("StockIncomeGet", "StockIncomeGet Start", DateTime.Now.ToString(), "", "");

            string result = null;
            int index = 0;

            string URL = "https://mops.twse.com.tw/nas/t21/sii/t21sc03_" + URL_YearMonth + ".csv";
            LogObj.WriteFileLog("StockIncomeGet", "StockIncomeGet Start", DateTime.Now.ToString(), URL, "");
            request = (HttpWebRequest)WebRequest.Create(URL);

            // 將 HttpWebRequest 的 Method 屬性設置為 GET
            request.Method = "GET";

            try
            {
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

                            string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
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
                            }
                        }
                        Thread.Sleep(100);
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("StockIncomeGet", e);
                return false;
            }
        }


        private string StockIncomeMessage(string StockID, string URL_YearMonth,string mYear , string season)
        {
            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("StockIncomeMessage", "StockIncomeMessage Start", DateTime.Now.ToString(), "", "");

            MSSQLCls SQLobj = new MSSQLCls();
            CommondFuncCls CfunObj = new CommondFuncCls();
            string incom_message = "";

            try
            {
                using (SqlConnection stockIncome_Conn = new SqlConnection(connstr))
                {
                    stockIncome_Conn.Open();
                    string sql = "select * from dbo.F_stockIncome_table('" + StockID + "', '" + URL_YearMonth + "')";
                    SQLobj.sqlDataTableCreater(dataset1, stockIncome_Conn, sql, "stockIncome");
                    if (dataset1.Tables["stockIncome"].Rows.Count == 0)
                    {
                        incom_message = "";
                    }
                    else
                    {
                        string StockName = dataset1.Tables["stockIncome"].Rows[0]["StockName"].ToString();
                        string DateMonth = dataset1.Tables["stockIncome"].Rows[0]["DateMonth"].ToString();
                        string InComeDiff_LastMonth = dataset1.Tables["stockIncome"].Rows[0]["InComeDiff_LastMonth"].ToString();
                        string InComeDiff_InCome_LastYear = dataset1.Tables["stockIncome"].Rows[0]["InComeDiff_InCome_LastYear"].ToString();
                        if (float.Parse(InComeDiff_LastMonth) > 0)
                        {
                            InComeDiff_LastMonth = "<FONT COLOR = RED>" + InComeDiff_LastMonth + "</FONT>";
                        }
                        if (float.Parse(InComeDiff_InCome_LastYear) > 0)
                        {
                            InComeDiff_InCome_LastYear = "<FONT COLOR = RED>" + InComeDiff_InCome_LastYear + "</FONT>";
                        }
                        incom_message = " ------- 營收比上月" + InComeDiff_LastMonth + "% 比去年" + InComeDiff_InCome_LastYear + "%";
                        string EPSsql = "select EPS from StockEPS where [StockID] = '" + StockID + "' and [EPS_Myear] = '" + mYear + "' and [EPS_season] = '" + season + "' ";
                        List<string> EPSLIST = SQLobj.sqlGetValue(EPSsql, stockIncome_Conn);
                        if (EPSLIST[0] != "_eof")
                        {
                            string epsstr = "";
                            if (float.Parse(EPSLIST[0].ToString()) > 0)
                            {
                                epsstr = "<FONT COLOR = RED>" + EPSLIST[0].ToString() + "</FONT>";
                            }

                            incom_message = incom_message + " ---EPS " + epsstr;
                        }
                    }
                }
                return incom_message;
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("StockIncomeMessage", e);
                return incom_message;
            }
        }



        public Boolean StockEPSGet(string myear, string season)
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
            string urlpath = "https://mops.twse.com.tw/mops/web/ajax_t163sb19?encodeURIComponent=1&step=1&firstin=1&TYPEK=sii&code=&year=" + myear + "&season=" + season;

            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("StockEPSGet", "StockEPSGet Start", DateTime.Now.ToString(), urlpath, "");

            try
            {   
                WebClient url = new WebClient();//3分鐘
                MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));
                HtmlDoc.Load(ms, Encoding.UTF8);

                CommondFuncCls CFNOBJ = new CommondFuncCls();

                HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes("//table[@class=\"hasBorder\"]");
                MSSQLCls DBObj = new MSSQLCls();
                string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
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
                                StockID = trNode.SelectSingleNode("td[1]").InnerText;
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
                return true;
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("StockEPSGet", e);
                return false;
            }
        }


        public Boolean StockMainTrendGet(string StockID)
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
            string urlpath = "https://www.wantgoo.com/stock/" + StockID + "/major-investors/main-trend";

            DebugCls LogObj = new DebugCls();
            LogObj.WriteFileLog("StockMainTrendGet", "StockMainTrendGet Start", DateTime.Now.ToString(), "", "");

            try
            {
                WebClient url = new WebClient();//3分鐘
                MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));
                HtmlDoc.Load(ms, Encoding.UTF8);

                CommondFuncCls CFNOBJ = new CommondFuncCls();

                HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes("//table[@id=\"main-trend\"]");
                MSSQLCls DBObj = new MSSQLCls();
                string connstr = "Password = dsc; User ID = sa; Initial Catalog = AutoStock; Data Source = 192.168.100.159; Connection Timeout = 0; Connection Lifetime = 4800; ";
                using (SqlConnection AutoStock_Conn = new SqlConnection(connstr))
                {
                    AutoStock_Conn.Open();
                    string Ins_sql = "";
                    foreach (var tableNode in htnode)
                    {
                        //CFNOBJ.showMessageBox("I am here");
                        var trNodes = tableNode.SelectNodes("./tr");

                       
                        string StockType = "";
                        string StockEPS = "";

                        foreach (var trNode in trNodes)
                        {
                            if (trNode.SelectSingleNode("td[1]") != null)
                            {
                                StockID = trNode.SelectSingleNode("td[1]").InnerText;
                                StockType = trNode.SelectSingleNode("td[3]").InnerText.Trim();
                                StockEPS = trNode.SelectSingleNode("td[4]").InnerText.Trim();
                                //string chksql = " select StockID from  StockEPS where StockID = '" + StockID + "' and EPS_Myear = '" + myear + "' and EPS_season = '" + season + "' ";
                               // List<string> chkList = DBObj.sqlGetValue(chksql, AutoStock_Conn);
                                //if (chkList[0] == "_eof")
                                //{
                                    //Ins_sql = " insert into StockEPS  (StockID,StockType,EPS_Myear,EPS_season,EPS) values ('" + StockID + "','" + StockType + "' , '" + myear + "','" + season + "','" + StockEPS + "' )";
                                    //DBObj.sqlExecute(Ins_sql, AutoStock_Conn);
                                //}
                            }
                        }
                    }

                }
                return true;
            }
            catch (Exception e)
            {
                LogObj.ErrHandle("StockEPSGet", e);
                return false;
            }
        }
        public class WebClientTo : WebClient
        {
            /// <summary>
            /// 过期时间
            /// </summary>
            public int Timeout { get; set; }

            public WebClientTo(int timeout)
            {
                Timeout = timeout;
            }

            /// <summary>
            /// 重写GetWebRequest,添加WebRequest对象超时时间
            /// </summary>
            /// <param name="address"></param>
            /// <returns></returns>
            protected override WebRequest GetWebRequest(Uri address)
            {
                HttpWebRequest request = (HttpWebRequest)base.GetWebRequest(address);
                request.Timeout = Timeout;
                request.ReadWriteTimeout = Timeout;
                return request;
            }
        }
    }

}
