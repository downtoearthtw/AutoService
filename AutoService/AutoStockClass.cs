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



namespace AutoStockService
{
    class AutoStockClass
    {
        private void DailyDatatoDB(string StockID,string sdt)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument HtmlDoc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument HtmlDoc1 = new HtmlAgilityPack.HtmlDocument();

            WebClient url = new WebClient();
            string urlpath = "";
            if (StockID == "0000")
            {
                urlpath = "https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date=" + sdt;
            }
            else
            {
                urlpath = "https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date=" + sdt + "&stockNo=" + StockID;
            }

            MemoryStream ms = new MemoryStream(url.DownloadData(urlpath));

            HtmlDoc.Load(ms, Encoding.UTF8);
            //HtmlDoc1.LoadHtml(HtmlDoc.DocumentNode.SelectSingleNode("/html/body/div/table/tbody/tr[11]").InnerHtml);

            if (StockID == "0000")
            {
                urlpath = "https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date=" + sdt;
            }
            else
            {
                HtmlNodeCollection htnode = HtmlDoc.DocumentNode.SelectNodes(@"/html/body/div/table/tbody/tr");
                MSSQLCls SQLobj = new MSSQLCls();

                using (SqlConnection Tons_Conn = new SqlConnection(LoginCls.TonsOfficeAConnStr))
                {
                    foreach (HtmlNode currTR in htnode)
                    {
                        string currTd1 = currTR.SelectSingleNode("td[1]").InnerText;
                        string currTd5 = currTR.SelectSingleNode("td[5]").InnerText;
                        string currTd6 = currTR.SelectSingleNode("td[6]").InnerText;
                        string currTd7 = currTR.SelectSingleNode("td[7]").InnerText;

                        string insSql = "insert into StockDailyData (StockID,DataDate,PriceH,PriceL) values ('" + StockID + "'," + currTd1 + "," + currTd5 + "," + currTd6 + "," + currTd7 + ")";
                    }
                }
            }
        }
    }


}
