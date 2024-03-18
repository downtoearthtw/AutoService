using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using ClassLibrary;
using System.Timers;
using AutoStockClass;
using System.Data.SqlClient;


namespace AutoService
{
    public partial class ServiceAutoStock : ServiceBase
    {

        public static int SecondInt = 0;
        public static int MinuteInt = 0;
        public static int minute5 = 0;

        public static int nowYearInt = 0;
        public static int nowMonthInt = 0;
        public static int nowDayInt = 0;
        public static int nowWeekInt = 0;
        public static int nowHourInt = 0;
        public static string nowWeekStr = "";
        public static Boolean startFlag = false;
        DateTime OldTime = new DateTime();
        string DailyStockDataIn_flag;
        string DailyKDCheck_flag;

        CommondFuncCls CfunObj = new CommondFuncCls();

        public ServiceAutoStock()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            DebugCls logObj = new DebugCls();
            logObj.WriteFileLog("ServiceJobExec", "OnStart", DateTime.Now.ToString() ,"", "");

            System.Timers.Timer timerClock = new System.Timers.Timer();
            timerClock.Elapsed += new ElapsedEventHandler(OnTimer);
            //timerClock.Interval = 360000; //六分鐘
            timerClock.Interval = (1 * 60 * 1000) * 3; // 3 min
                                                       // 1 min

            timerClock.Enabled = true;
            OldTime = DateTime.Now;
            DailyStockDataIn_flag = "";
            DailyKDCheck_flag = "";
            DebugCls LogObj = new DebugCls();
        }

        protected override void OnStop()
        {
            DebugCls logObj = new DebugCls();
            logObj.WriteFileLog("ServiceJobExec", "OnStop", DateTime.Now.ToString(), "", "");
        }

        string working_flag = "";

        public void OnTimer(Object source, ElapsedEventArgs e)
        {
            DebugCls logObj = new DebugCls();
            logObj.WriteFileLog("ServiceJobExec", "OnTimer", DateTime.Now.ToString(), working_flag, "");

            StockDataProcess STDObj = new StockDataProcess();
            DateTime CurrentTime = DateTime.Now;
            string CurrentTimeLen8 = CfunObj.DateToStringLen8(CurrentTime);
            nowHourInt = CurrentTime.Hour;
            if (nowHourInt == 12) DailyStockDataIn_flag = ""; //中午的時候把這個flag清掉，讓下午可以再做一次
            int oldHourInt = OldTime.Hour;

            logObj.WriteFileLog("ServiceJobExec", "OnTimer", DateTime.Now.ToString(), CurrentTime.DayOfWeek.ToString(), "DayofWeek");

            if (nowHourInt > 1 && nowHourInt < 4)//改成半夜執行
            {
                if (working_flag != "working")
                {
                    if (DailyStockDataIn_flag != CurrentTimeLen8) //每日資料採集DailyStockDataIn_flag，如果不是今天的8碼日期，表示還沒做，周末不用採集資料
                    {
                        working_flag = "working";

                        Boolean isdone = STDObj.DailyStockDataIn(); //資料採集DailyStockDataIn
                        logObj.WriteFileLog("ServiceJobExec", "DailyStockDataIn", DateTime.Now.ToString(), CurrentTime.DayOfWeek.ToString(), isdone.ToString());

                        DateTime everyMonth = DateTime.Now.AddMonths(-1);
                            
                        string mYear = (everyMonth.Year - 1911).ToString();
                        string mMonth = (everyMonth.Month).ToString();
                        string URL_YearMonth = mYear + "_" + mMonth;
                        string season = "";
                        if (mMonth == "1" || mMonth == "2" || mMonth == "3")
                        {
                            mYear = (int.Parse(mYear) - 1).ToString();
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

                        string StockIncomeGetFlagError = "ERROR";
                        while (StockIncomeGetFlagError == "ERROR")
                        {
                            try
                            {
                                STDObj.StockIncomeGet(URL_YearMonth); //取得上個月的營業收入
                                StockIncomeGetFlagError = "PASS";
                            }
                            catch
                            {
                                StockIncomeGetFlagError = "ERROR";
                            }
                        }

                        logObj.WriteFileLog("ServiceJobExec", "StockIncomeGet", DateTime.Now.ToString(), CurrentTime.DayOfWeek.ToString(), isdone.ToString());

                        string StockEPSGet = "ERROR";
                        while (StockEPSGet == "ERROR")
                        {
                            try
                            {
                                STDObj.StockEPSGet(mYear, season); //取得上一季的EPS
                                StockEPSGet = "PASS";
                            }
                            catch
                            {
                                StockEPSGet = "ERROR";
                            }
                        }
                        logObj.WriteFileLog("ServiceJobExec", "StockIncomeGet", DateTime.Now.ToString(), CurrentTime.DayOfWeek.ToString(), isdone.ToString());

                        isdone = true;
                        if (isdone)
                        {
                            DailyStockDataIn_flag = CurrentTimeLen8;
                        }
                        else
                        {
                            working_flag = "not working";
                        }
                    }
                    working_flag = "not working";
                }

                if (working_flag != "working")
                {
                    if (DailyKDCheck_flag != CurrentTimeLen8) //每日KD檢查DailyKDCheck_flag，如果不是今天的8碼日期，表示還沒做
                    {
                        working_flag = "working";
                        if (CurrentTime.DayOfWeek.ToString() != "Sunday" )
                        {
                            logObj.WriteFileLog("ServiceJobExec", "DailyKDCheck", DateTime.Now.ToString(), CurrentTime.DayOfWeek.ToString(), "DayofWeek");
                            Boolean isdone = false;
                            isdone = STDObj.DailyKDCheck();
                            if (isdone)
                            {
                                DailyKDCheck_flag = CurrentTimeLen8;
                            }
                            else
                            {
                                working_flag = "not working";
                            }
                        }
                        working_flag = "not working";
                    }
                }
            }
            OldTime = CurrentTime;
        }
    }
}
