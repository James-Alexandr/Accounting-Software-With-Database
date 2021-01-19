using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    class tableAdapterr
    {
     //Journal table
        public int ID{ get; set; }
        public string Date { get; set; }
        public string Account_1 { get; set; }
        public string Account_2 { get; set; }
        public string Type_1 { get; set; }
        public string Type_2 { get; set; }
        public double Debit { get; set; }
        public double Credit { get; set; }

        //T Accounts
        public int ID_TAcccounts { get; set; }
        public string Account_TAccounts{ get; set; }
        public string Type_TAccounts { get; set; }
        public string DebitList { get; set; }
        public string CreditList { get; set; }
        public double TotalDebit { get; set; }
        public double TotalCredit { get; set; }
        public double Balance { get; set; }

        //Balance Sheet
        public int ID_BalanceSheet { get; set; }
        public string Asset_Account_Name { get; set; }
        public double Asset_Account_Balance { get; set; }
        public double Total_Assets { get; set; }
        public string Liability_Account_Name { get; set; }
        public double Liability_Account_Balance { get; set; }
        public double Total_Liability { get; set; }
        

        //Income Statement
        public int ID_IncomeStatement { get; set; }
        public string Revenue_Account_Name { get; set; }
        public double Revenue_Account_Balance { get; set; }
        public double Total_Revenue { get; set; }
        public string Expense_Account_Name { get; set; }
        public double Expense_Account_Balance { get; set; }
        public double Total_Expense { get; set; }
        public double Net_Income { get; set; }
    
        //Statement of Owner Equity
        public int ID_StatementOfOE { get; set; }
        public double Start_Capital { get; set; }
        public double Net_Income_StatementOfOE { get; set; }
        public double Total_Withdrawals { get; set; }
        public double FInal_Capital { get; set; }
    
    }
}
