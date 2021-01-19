using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog

{
    /// <summary>
    /// This class allows the user to create new tables into the DB
    /// </summary>
    class AddDBTable
    {

        //****************T Accounts *******************

        public static void TAccounts (string tableName)
        {
            LeapLogDBManager sqlTables = new LeapLogDBManager();

            string TaccountName = tableName + "Taccount";

            string dbString = @"CREATE TABLE  " + TaccountName + "( ID INT IDENTITY(1, 1) NOT NULL, Account  NVARCHAR(50) NULL, Type NVARCHAR(50) NULL," +
    "DebitList NVARCHAR(50) NULL, CreditList NVARCHAR(50) NULL, " +
    "TotalDebit FLOAT NULL, TotalCredit  FLOAT NULL, Balance FLOAT NULL, PRIMARY KEY CLUSTERED(Id ASC))";
            sqlTables.WriteData(dbString);
        }


        //****************Balance sheet*******************
        public static void BalanceSheet(string tableName)
        {
            LeapLogDBManager sqlTables = new LeapLogDBManager();

            string BalanceSheetName = tableName + "BalanceSheet";

            string dbString = @"CREATE TABLE  " + BalanceSheetName + "( ID INT IDENTITY(1, 1) NOT NULL, Asset_Account_Name  NVARCHAR(50) NULL,  Asset_Account_Balance  FLOAT NULL, Total_Assets FLOAT NULL," +
    "Liability_Account_Name NVARCHAR(50) NULL, Liability_Account_Balance FLOAT NULL, " +
    "Total_Liability FLOAT NULL,PRIMARY KEY CLUSTERED(Id ASC))";
            sqlTables.WriteData(dbString);
        }


        //****************Income Statement*******************
        public static void IncomeStatement(string tableName)
        {
            LeapLogDBManager sqlTables = new LeapLogDBManager();
            string IncomeStatementName = tableName + "IncomeStatement";


            string dbString = @"CREATE TABLE  " + IncomeStatementName + "( ID INT IDENTITY(1, 1) NOT NULL, Revenue_Account_Name NVARCHAR(50) NULL, Revenue_Account_Balance FLOAT NULL, Total_Revenue FLOAT NULL, Expense_Account_Name NVARCHAR(50) NULL, Expense_Account_Balance FLOAT NULL, Total_Expense FLOAT NULL," +
    "Net_Income FLOAT NULL,PRIMARY KEY CLUSTERED(Id ASC))";
            sqlTables.WriteData(dbString);
        }


        //****************Statement of Owner Equity*******************
        public static void StatementOfOwnerEquity(string tableName)
        {
            LeapLogDBManager sqlTables = new LeapLogDBManager();
            string StatementOfOEName = tableName + "StatementOfOE";


            string dbString = @"CREATE TABLE  " + StatementOfOEName + "( ID INT IDENTITY(1, 1) NOT NULL,Start_Capital FLOAT NULL, Net_Income FLOAT NULL, Total_Withdrawals FLOAT NULL," +
    "Final_Capital FLOAT NULL,PRIMARY KEY CLUSTERED(Id ASC))";
            sqlTables.WriteData(dbString);
        }

         



    }

}
