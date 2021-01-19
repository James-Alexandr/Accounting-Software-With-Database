using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Controls;

namespace LeapLog
{
    //The database class stores all the t-accounts created by the user
     public static class Database
     {

        //t-account data
        private static List<Entry_tacc> tentries;
        internal static List<Entry_tacc> TEntries { get => tentries; set => tentries = value; }

        //income statement data
        private static income_data incomeData;
        internal static income_data IncomeData { get => incomeData; set => incomeData = value; }

        //soe data
        private static soe_data soeData;
        internal static soe_data SoeData { get => soeData; set => soeData = value; }

        //balance sheet data
        private static balance_data balanceData;
        internal static balance_data BalanceData { get => balanceData; set => balanceData = value; }

        // select date for balance sheet calendar
        public static DateTime select_date;
          // select date range for income-statement
          public static DateRange is_date;
          // select date range for statement of owner's equity
          public static DateRange sooe_date;
          // net-income
          public static double net_income = 0;

        static Database()
        {
            tentries = new List<Entry_tacc>();

            incomeData = new income_data();

            soeData = new soe_data();

            balanceData = new balance_data();
        }
    }
}
