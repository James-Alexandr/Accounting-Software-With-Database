using Microsoft.Office.Interop.Excel;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
 



namespace LeapLog
{
    /// <summary>
    /// The Journal class allows the user to create new journal entries that can be stored
    /// using the Database class.
    /// </summary>
    public partial class Journal : UserControl
    {
        public static string passingText;




        public Journal()
        {
            InitializeComponent();
        }
        
        //Add new entry button
        private void addEntryBtn_Click(object sender, RoutedEventArgs e)
        {
            //new entry is created
            Entry tempEntry = new Entry();

            //clear any warnings, if necessary
            warningTB.Visibility = Visibility.Hidden;
            warningAT.Visibility = Visibility.Hidden;
            warningAN.Visibility = Visibility.Hidden;

            //populates new entry object with user data given
            try
            {
                //process account names to make sure no quotation marks are entered
                if (account1TB.Text.Contains("'") || account2TB.Text.Contains("'"))
                {
                    throw new Exception("quotation");
                }

                // Check for account types.
                if (type1CB.SelectedItem == null || type2CB.SelectedItem == null)
                {

                    throw new Exception("null_account_types");
                }

                // Check for blank account names.
                if (account1TB.Text.Replace(" ", "") == string.Empty || account2TB.Text.Replace(" ", "") == string.Empty)
                {
                    throw new Exception("blank_account_names");
                }

                tempEntry.Account1 = account1TB.Text;
                tempEntry.Account2 = "      " + account2TB.Text;
                tempEntry.Debit = Double.Parse(debitTB.Text);
                tempEntry.Credit = Double.Parse(creditTB.Text);
                tempEntry.Type1 = type1CB.Text;
                tempEntry.Type2 = type2CB.Text;

                //add entry into entryGrid
                entryGrid.Items.Add(tempEntry);

                //separate entry into two t-accounts
                List<Entry_tacc> tempAccounts = T_Accounts.get_taccs(tempEntry);
                //process t-accounts
                T_Accounts.add_taccs(tempAccounts);

                //<<-------------inserting data to the database-------------------->>
                LeapLogDBManager sqlTables = new LeapLogDBManager();
                 
                //<<------- this chooses the table where the data will be added to-------->>
                 string tableName = user_Input.Text.Replace(" ", "");
 
                //**********insert into journal table************
                sqlTables.WriteData("INSERT INTO " + tableName + " VALUES ('" + DateTime.Now + "','" + account1TB.Text + "','" + account2TB.Text + "','" + type1CB.Text + "','" + type2CB.Text + "','" + double.Parse(debitTB.Text) + "','" + double.Parse(creditTB.Text) + "')");

            }
            catch (Exception error){
                switch (error.Message)
                {
                    case "quotation":
                        //if incorrect data entered, warning given
                        //and entry not saved
                        warningTB.Visibility = Visibility.Visible;
                        break;
                    case "null_account_types":
                        // if the account types are null, display a warning.
                        warningAT.Visibility = Visibility.Visible;
                        break;
                    case "blank_account_names":
                        // if account names are blank display warning.
                        warningAN.Visibility = Visibility.Visible;
                        break;
                }
            }

            //clear textboxes
            account1TB.Text = "";
            account2TB.Text = "";
            debitTB.Text = "";
            creditTB.Text = "";
            type1CB.SelectedItem = null;
            type2CB.SelectedItem = null;

            account1TB.Focus();

        }

        private void enter_button_Click(object sender, RoutedEventArgs e)
        {

            string messageBoxText = "Database field cannot be null or empty";
            string caption = "Wrong Input";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;
            //**********************Check if table name already exists**********************

            List<String> list = new List<String>(); // to save the table names

            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            int counterV = 0;
            string sql = null;

            connetionString = $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True";
            sql = "Select DISTINCT(name) FROM sys.Tables";

            connection = new SqlConnection(connetionString);

            try
            {
                connection.Open();
                command = new SqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                connection.Close();

                for (counterV = 0; counterV <= ds.Tables[0].Rows.Count - 1; counterV++)
                {
                    list.Add(ds.Tables[0].Rows[counterV].ItemArray[0].ToString());
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Can not open connection ! ");
            }
            //******************end of block of code*********************************


            //******************Control if statement******************

            if (list.FindIndex(x => x.Equals(user_Input.Text.Trim().Replace(" ", ""), StringComparison.CurrentCultureIgnoreCase)) != -1)

                MessageBox.Show("Database name already taken.", "Try again", button, icon);

            else
            {
                string tableName = user_Input.Text.Replace(" ", "");
                // If the tableName is empty quit this function.
                if (tableName.Length < 1)
                {
                    MessageBox.Show("No name was given to the table.", "Try again", button, icon);
                    return;
                }
                // Fail safe for if the user inputs an number as the first word.
                Match match = Regex.Match(tableName[0].ToString(), "[0-9]");
                if (match.Success)
                {
                    MessageBox.Show("Letters only as the first character.", "Try again", button, icon);
                    return;
                }
                //<<--------this creates the datatable into the database------->>
                LeapLogDBManager sqlTables = new LeapLogDBManager();
                //string journalName = tableName + "Journal";
                string dbString = @"CREATE TABLE  " + tableName + "( ID INT IDENTITY(1, 1) NOT NULL,Date DATE NULL, Account_1  NVARCHAR(50) NULL, Account_2 NVARCHAR(50) NULL," +
    "Type_1 NVARCHAR(50) NULL, Type_2 NVARCHAR(50) NULL, " +
    "Debit MONEY NULL, Credit  MONEY NULL,PRIMARY KEY CLUSTERED(Id ASC))";
                if (String.IsNullOrEmpty(user_Input.Text) || user_Input.Text == "")
                {
                    MessageBox.Show(messageBoxText, caption, button, icon);
                }


                else
                {
                    sqlTables.WriteData(dbString);
                    MessageBox.Show("Tables " + user_Input.Text + " added to Database. ", "Table added", button, icon);

                    AddDBTable.TAccounts(tableName);
                    AddDBTable.BalanceSheet(tableName);
                    AddDBTable.IncomeStatement(tableName);
                    AddDBTable.StatementOfOwnerEquity(tableName);
                  


                }
            }
        }

        //button that opens up journal help feature
        private void journalHelpButton_Click(object sender, RoutedEventArgs e)
        {
            if (journalHelpWindow1.Visibility == Visibility.Collapsed &&
                journalHelpWindow2.Visibility == Visibility.Collapsed &&
                journalHelpWindow3.Visibility == Visibility.Collapsed)
            {
                journalHelpWindow1.Visibility = Visibility.Visible;
            }
            else
            {
                journalHelpWindow1.Visibility = Visibility.Collapsed;
                journalHelpWindow2.Visibility = Visibility.Collapsed;
                journalHelpWindow3.Visibility = Visibility.Collapsed;
            }
        }

        //button that changes to second help feature page
        private void forwardButton_Click(object sender, RoutedEventArgs e)
        {
            journalHelpWindow2.Visibility = Visibility.Visible;
            journalHelpWindow1.Visibility = Visibility.Collapsed;
            journalHelpWindow3.Visibility = Visibility.Collapsed;
        }

        //button that changes to first help feature page
        private void backButton_Click(object sender, RoutedEventArgs e)
        {
            journalHelpWindow1.Visibility = Visibility.Visible;
            journalHelpWindow2.Visibility = Visibility.Collapsed;
            journalHelpWindow3.Visibility = Visibility.Collapsed;
        }

        //button that changes to third help window
        private void forwardButton2_Click(object sender, RoutedEventArgs e)
        {
            journalHelpWindow3.Visibility = Visibility.Visible;
            journalHelpWindow1.Visibility = Visibility.Collapsed;
            journalHelpWindow2.Visibility = Visibility.Collapsed;
        }

        //button that changes from third to second help window
        private void backButton2_Click(object sender, RoutedEventArgs e)
        {
            journalHelpWindow3.Visibility = Visibility.Collapsed;
            journalHelpWindow2.Visibility = Visibility.Visible;
            journalHelpWindow1.Visibility = Visibility.Collapsed;
        }

       
        
        //************to excel *********************


        private void toExcel_Click(object sender, RoutedEventArgs e)
        {
            if (Save_DB() == false) return;
            //******************check if table is in the db **********************

            MessageBoxButton button2 = MessageBoxButton.OK;
            MessageBoxImage icon2 = MessageBoxImage.Warning;
            //**********************Check if table name already exists**********************

            List<String> list = new List<String>(); // to save the table names

            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            int counterV = 0;
            string sql = null;

            connetionString = $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True";
            sql = "Select DISTINCT(name) FROM sys.Tables";

            connection = new SqlConnection(connetionString);

            try
            {
                connection.Open();
                command = new SqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                connection.Close();

                for (counterV = 0; counterV <= ds.Tables[0].Rows.Count - 1; counterV++)
                {
                    list.Add(ds.Tables[0].Rows[counterV].ItemArray[0].ToString());
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Can not open connection ! ");
            }
            //******************end of block of code*********************************


            //******************Control if statement****************** MessageBox.Show("Table name already taken.", "Try again", button, icon);

            if (list.FindIndex(x => x.Equals(user_Input.Text.Trim().Replace(" ", ""), StringComparison.CurrentCultureIgnoreCase)) != -1)

            {  


            //******************check if table is in the db **********************

            // Save the database before exporting.
            

            string messageBoxText = "Journal field cannot be null or empty";
            string caption = "Wrong Input";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;

            string tableName = user_Input.Text.Replace(" ", "");

            string TaccountName = tableName + "Taccount";
            string BalanceSheetName = tableName + "BalanceSheet";
            string IncomeStatementName = tableName + "IncomeStatement";
            string StatementOfOEName = tableName + "StatementOfOE";

            System.Diagnostics.Debug.WriteLine(BalanceSheetName);

            List<tableAdapterr> DBList = new List<tableAdapterr>();



                //**************************flow control************************************
                if (String.IsNullOrEmpty(user_Input.Text) || user_Input.Text == "")
                {
                    MessageBox.Show(messageBoxText, caption, button, icon);
                }

                //***********************from db to table adapter*******************************        

                else
                {
                    //******************************************************************************************************
                    //*************extracting journal table********************
                    //******************************************************************************************************
                    SqlConnection conn = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True");
                    string query = "Select * from " + tableName + " ";
                    SqlDataAdapter sda = new SqlDataAdapter(query, conn);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    sda.Fill(dataTable);
                    //string name;


                    //*********************if need assistance please ask James***************************



                    foreach (DataRow row in dataTable.Rows)
                    {
                        tableAdapterr journalTable = new tableAdapterr();

                        /*name = row["username"].ToString().Trim();
                       System.Diagnostics.Debug.WriteLine(name);
                       list.Add(name);*/
                        journalTable.ID = Convert.ToInt32(row["ID"]);
                        journalTable.Date = row["Date"].ToString().Trim();
                        journalTable.Account_1 = row["Account_1"].ToString().Trim();
                        journalTable.Account_2 = row["Account_2"].ToString().Trim();
                        journalTable.Type_1 = row["Type_1"].ToString().Trim();
                        journalTable.Type_2 = row["Type_2"].ToString().Trim();
                        journalTable.Debit = Convert.ToDouble(row["Debit"]);
                        journalTable.Credit = Convert.ToDouble(row["Credit"]);

                        //*********adding to list adapter*********
                        DBList.Add(journalTable);
                    }

                    //test code
                    foreach (var i in DBList)
                    {
                        System.Diagnostics.Debug.WriteLine(i.Account_1);
                    }



                    //******************************************************************************************************
                    //*************extracting T Accounts table********************
                    //******************************************************************************************************
                    SqlConnection conn_TA = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True");
                    string query_TA = "Select * from " + TaccountName + " ";
                    SqlDataAdapter sda_TA = new SqlDataAdapter(query_TA, conn_TA);
                    System.Data.DataTable dataTable_TA = new System.Data.DataTable();
                    sda_TA.Fill(dataTable_TA);
                    //string name;


                    //*********************if need assistance please ask James***************************



                    //***********************Extraction journal data from DB to journalTable then to DBList adapter*******************************        

                    foreach (DataRow row in dataTable_TA.Rows)
                    {
                        tableAdapterr TAccountTable = new tableAdapterr();


                        TAccountTable.ID_TAcccounts = Convert.ToInt32(row["ID"]);
                        TAccountTable.Account_TAccounts = row["Account"].ToString().Trim();
                        TAccountTable.Type_TAccounts = row["Type"].ToString().Trim();
                        TAccountTable.DebitList = row["DebitList"].ToString();
                        TAccountTable.CreditList = row["CreditList"].ToString();
                        TAccountTable.TotalDebit = Convert.ToDouble(row["TotalDebit"]);
                        TAccountTable.TotalCredit = Convert.ToDouble(row["TotalCredit"]);
                        TAccountTable.Balance = Convert.ToDouble(row["Balance"]);


                        //*********adding to list adapter*********
                        DBList.Add(TAccountTable);
                    }


                    //******************************************************************************************************
                    //*************extracting Balance Sheet********************
                    //******************************************************************************************************
                    SqlConnection conn_BS = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True");
                    string query_BS = "Select * from " + BalanceSheetName + " ";
                    SqlDataAdapter sda_BS = new SqlDataAdapter(query_BS, conn_BS);
                    System.Data.DataTable dataTable_BS = new System.Data.DataTable();
                    sda_BS.Fill(dataTable_BS);
                    //string name;


                    //*********************if need assistance please ask James***************************



                    //***********************Extraction journal data from DB to journalTable then to DBList adapter*******************************        

                    foreach (DataRow row in dataTable_BS.Rows)
                    {
                        tableAdapterr BalanceSheetTable = new tableAdapterr();


                        BalanceSheetTable.ID_BalanceSheet = Convert.ToInt32(row["ID"]);
                        BalanceSheetTable.Asset_Account_Name = row["Asset_Account_Name"].ToString().Trim();
                        BalanceSheetTable.Asset_Account_Balance = Convert.ToDouble(row["Asset_Account_Balance"]);
                        BalanceSheetTable.Total_Assets = Convert.ToDouble(row["Total_Assets"]);
                        BalanceSheetTable.Liability_Account_Name = row["Liability_Account_Name"].ToString().Trim();
                        BalanceSheetTable.Liability_Account_Balance = Convert.ToDouble(row["Liability_Account_Balance"]);
                        BalanceSheetTable.Total_Liability = Convert.ToDouble(row["Total_Liability"]);


                        //*********adding to list adapter*********
                        DBList.Add(BalanceSheetTable);
                    }



                    //******************************************************************************************************
                    //*************extracting Income Statement********************
                    //******************************************************************************************************
                    SqlConnection conn_IS = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True");
                    string query_IS = "Select * from " + IncomeStatementName + " ";
                    SqlDataAdapter sda_IS = new SqlDataAdapter(query_IS, conn_IS);
                    System.Data.DataTable dataTable_IS = new System.Data.DataTable();
                    sda_IS.Fill(dataTable_IS);
                    //string name;

                    //***********************Extraction journal data from DB to journalTable then to DBList adapter*******************************        

                    foreach (DataRow row in dataTable_IS.Rows)
                    {
                        tableAdapterr IncomeStatementTable = new tableAdapterr();


                        IncomeStatementTable.ID_IncomeStatement = Convert.ToInt32(row["ID"]);
                        IncomeStatementTable.Revenue_Account_Name = row["Revenue_Account_Name"].ToString().Trim();
                        IncomeStatementTable.Revenue_Account_Balance = Convert.ToDouble(row["Revenue_Account_Balance"]);
                        IncomeStatementTable.Total_Revenue = Convert.ToDouble(row["Total_Revenue"]);
                        IncomeStatementTable.Expense_Account_Name = row["Expense_Account_Name"].ToString().Trim();
                        IncomeStatementTable.Expense_Account_Balance = Convert.ToDouble(row["Expense_Account_Balance"]);
                        IncomeStatementTable.Total_Expense = Convert.ToDouble(row["Total_Expense"]);
                        IncomeStatementTable.Net_Income = Convert.ToDouble(row["Net_Income"]);


                        //*********adding to list adapter*********
                        DBList.Add(IncomeStatementTable);
                    }


                    //******************************************************************************************************
                    //*************extracting StatementOfOEName********************
                    //******************************************************************************************************
                    SqlConnection conn_SOE = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\Database1.MDF;Integrated Security=True");
                    string query_SOE = "Select * from " + StatementOfOEName + " ";
                    SqlDataAdapter sda_SOE = new SqlDataAdapter(query_SOE, conn_SOE);
                    System.Data.DataTable dataTable_SOE = new System.Data.DataTable();
                    sda_SOE.Fill(dataTable_SOE);
                    //string name;

                    //***********************Extraction journal data from DB to journalTable then to DBList adapter*******************************        

                    foreach (DataRow row in dataTable_SOE.Rows)
                    {
                        tableAdapterr OwnerEquityTable = new tableAdapterr();


                        OwnerEquityTable.ID_StatementOfOE = Convert.ToInt32(row["ID"]);
                        OwnerEquityTable.Start_Capital = Convert.ToDouble(row["Start_Capital"]);
                        OwnerEquityTable.Net_Income_StatementOfOE = Convert.ToDouble(row["Net_Income"]);
                        OwnerEquityTable.Total_Withdrawals = Convert.ToDouble(row["Total_Withdrawals"]);
                        OwnerEquityTable.FInal_Capital = Convert.ToDouble(row["Final_Capital"]);


                        //*********adding to list adapter*********
                        DBList.Add(OwnerEquityTable);
                    }

                    //******************************************************************************************************
                    //------------------------------------to excel------------------------------------------
                    //******************************************************************************************************



                    //*********************if need assistance please ask ***************************
                    //*****************************James Alexander **************************

                    // Load up Excel, then make a new empty workbook.
                    Excel.Application excelApp = new Excel.Application();



                    object missing = System.Reflection.Missing.Value;

                    Excel.Workbook oWB = excelApp.Workbooks.Add(missing);

                    //excelApp.Workbooks.Add();
                    // This example uses a single workSheet.
                    Worksheet workSheet = (Worksheet)oWB.ActiveSheet;
                    workSheet.Name = "Journal";
                    // Establish column headings in cells.
                    workSheet.Cells[1, "A"] = "ID";
                    workSheet.Cells[1, "B"] = "Date";
                    workSheet.Cells[1, "C"] = "Account 1";
                    workSheet.Cells[1, "D"] = "Type 1";
                    workSheet.Cells[1, "E"] = "Debit";
                    workSheet.Cells[1, "F"] = "Account 2";
                    workSheet.Cells[1, "G"] = "Type 2";
                    workSheet.Cells[1, "H"] = "Credit";

                    // Now, map all data in List<tableAdapterr> to the cells of the spreadsheet.
                    //****************T Accounts *******************
                    Worksheet workSheet2 = (Worksheet)oWB.Sheets.Add(missing, missing, 1, missing);
                    workSheet2.Name = "T Accounts";
                    // Establish column headings in cells.
                    workSheet2.Cells[1, "A"] = "ID";
                    workSheet2.Cells[1, "B"] = "Account";
                    workSheet2.Cells[1, "C"] = "Type";
                    workSheet2.Cells[1, "D"] = "List of Debits";
                    workSheet2.Cells[1, "E"] = "List of Credits";
                    workSheet2.Cells[1, "F"] = "Total Debit";
                    workSheet2.Cells[1, "G"] = "Total Credit";
                    workSheet2.Cells[1, "H"] = "Balance";

                    //****************Balance sheet*******************
                    Worksheet workSheet3 = (Worksheet)oWB.Sheets.Add(missing, missing, 1, missing);
                    workSheet3.Name = "Balance Sheet";
                    // Establish column headings in cells.
                    workSheet3.Cells[1, "A"] = "ID";
                    workSheet3.Cells[1, "B"] = "Asset Account Name";
                    workSheet3.Cells[1, "C"] = "Asset Account Balance";
                    workSheet3.Cells[1, "D"] = "Total Assets";
                    workSheet3.Cells[1, "E"] = "Liability or Owners Equity Account Name";
                    workSheet3.Cells[1, "F"] = "Liability or Owners Equity Account Balance";
                    workSheet3.Cells[1, "G"] = "Total Liabilities and Owners Equity";

                    //****************Income Statement*******************
                    Worksheet workSheet4 = (Worksheet)oWB.Sheets.Add(missing, missing, 1, missing);
                    workSheet4.Name = "Income Statement";
                    // Establish column headings in cells.
                    workSheet4.Cells[1, "A"] = "ID";
                    workSheet4.Cells[1, "B"] = "Revenue Account Name";
                    workSheet4.Cells[1, "C"] = "Revenue Account Balance";
                    workSheet4.Cells[1, "D"] = "Total Revenue";
                    workSheet4.Cells[1, "E"] = "Expense Account Name";
                    workSheet4.Cells[1, "F"] = "Expense Account Balance";
                    workSheet4.Cells[1, "G"] = "Total Expense";
                    workSheet4.Cells[1, "H"] = "Net Income";

                    //****************Statement of Owner Equity*******************
                    Worksheet workSheet5 = (Worksheet)oWB.Sheets.Add(missing, missing, 1, missing);
                    workSheet5.Name = "Statement of Owner Equity";
                    // Establish column headings in cells.
                    workSheet5.Cells[1, "A"] = "ID";
                    workSheet5.Cells[1, "B"] = "Start Capital";
                    workSheet5.Cells[1, "C"] = "Net Income";
                    workSheet5.Cells[1, "D"] = "Total Withdrawals";
                    workSheet5.Cells[1, "E"] = "Final Capital";




                    // Now, map all data in List<tableAdapterr> to the cells of the spreadsheet (journal).
                    int row1 = 1;

                    foreach (var i in DBList)
                    {
                        if (i.Debit != 0)
                        {
                            row1++;

                            //workSheet.Cells[row1, "A"] = i.ID;
                            workSheet.Cells[row1, "B"] = i.Date;
                            workSheet.Cells[row1, "C"] = i.Account_1;
                            workSheet.Cells[row1, "D"] = i.Type_1;
                            workSheet.Cells[row1, "E"] = i.Debit;
                            workSheet.Cells[row1, "F"] = i.Account_2;
                            workSheet.Cells[row1, "G"] = i.Type_2;
                            workSheet.Cells[row1, "H"] = i.Credit;
                        }
                    }


                    // Now, map all data in List<tableAdapterr> to the cells of the Sheet 2 (t-accounts).

                    int taccRow = 1;

                    //for every account, pull data
                    foreach (var i in DBList)
                    {
                        
                        if (i.DebitList != null)
                        {
                            taccRow++;
                            // workSheet2.Cells[taccRow, "A"] = i.ID_TAcccounts;    
                            workSheet2.Cells[taccRow, "B"] = i.Account_TAccounts;
                            workSheet2.Cells[taccRow, "C"] = i.Type_TAccounts;
                            workSheet2.Cells[taccRow, "D"] = i.DebitList;
                            workSheet2.Cells[taccRow, "E"] = i.CreditList;
                            workSheet2.Cells[taccRow, "F"] = i.TotalDebit;
                            workSheet2.Cells[taccRow, "G"] = i.TotalCredit;
                            workSheet2.Cells[taccRow, "H"] = i.Balance;
                        }

                    }

                    // Now, map all data in List<tableAdapterr> to the cells of the Sheet 3 (balance sheet).

                    int assetRow = 1;
                    int loeRow = 1;
                    int totalRow = 1;
                    //iterate through list of all asset accounts
                    foreach (var i in DBList)
                    {


                        //get the account name and add to excel sheet
                        if (i.Asset_Account_Name != null && i.Asset_Account_Balance != 0)
                        {
                            assetRow++;
                            workSheet3.Cells[assetRow, "B"] = i.Asset_Account_Name;
                            workSheet3.Cells[assetRow, "C"] = i.Asset_Account_Balance;
                        }
                        if (i.Total_Assets != 0)
                        {
                            totalRow++;
                            workSheet3.Cells[2, "D"] = i.Total_Assets;
                        }
                    }



                    //iterate through list of all other accounts
                    foreach (var i in DBList)
                    {


                        //get the account name and add to excel sheet
                        if (i.Liability_Account_Name != null && i.Liability_Account_Balance != 0)
                        {
                            loeRow++;
                            workSheet3.Cells[loeRow, "E"] = i.Liability_Account_Name;
                            workSheet3.Cells[loeRow, "F"] = i.Liability_Account_Balance;
                        }
                        if (i.Total_Liability != 0)
                        {
                            totalRow++;
                            workSheet3.Cells[2, "G"] = i.Total_Liability;
                        }

                    }


                    // Now, map all data in List<tableAdapterr> to the cells of the Sheet 4 (income statement).

                    //Predicate<tableAdapterr> totalRevenueFinder = (tableAdapterr t) => { return t.Total_Revenue != 0; };
                    int totalsRow = 2;
                    int revRow = 1;
                    int expRow = 1;
                    foreach (var i in DBList)

                    {

                        // tableAdapterr finder = DBList.Find(totalRevenueFinder);
                        if (i.Net_Income != 0)
                        {
                            workSheet4.Cells[totalsRow, "H"] = i.Net_Income;
                            workSheet4.Cells[totalsRow, "D"] = i.Total_Revenue;
                            workSheet4.Cells[totalsRow, "G"] = i.Total_Expense;
                        }

                        if (i.Revenue_Account_Name != null && i.Revenue_Account_Balance != 0)
                        {
                            revRow++;
                            workSheet4.Cells[revRow, "B"] = i.Revenue_Account_Name;
                            workSheet4.Cells[revRow, "C"] = i.Revenue_Account_Balance;
                        }

                        if (i.Expense_Account_Name != null && i.Expense_Account_Balance != 0)
                        {
                            expRow++;
                            workSheet4.Cells[expRow, "E"] = i.Expense_Account_Name;
                            workSheet4.Cells[expRow, "F"] = i.Expense_Account_Balance;
                        }

                    }

                    // Now, map all data in List<tableAdapterr> to the cells of the Sheet 5 (oe statement).

                    int SOERow = 1;
                    foreach (var i in DBList)
                    {

                        
                        if (i.Start_Capital != 0)
                        {
                            SOERow++;
                            workSheet5.Cells[SOERow, "B"] = i.Start_Capital;
                            workSheet5.Cells[SOERow, "C"] = i.Net_Income_StatementOfOE;
                            workSheet5.Cells[SOERow, "D"] = i.Total_Withdrawals;
                            workSheet5.Cells[SOERow, "E"] = i.FInal_Capital;
                        }
                    }

                    ;

                    // Give our table data a nice look and feel.
                    workSheet.Range["A1"].AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet.Columns.EntireColumn.AutoFit();
                    workSheet.Range["A1", "H1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;

                    workSheet2.Range["A1"].AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet2.Columns.EntireColumn.AutoFit();
                    workSheet2.Range["A1", "H1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;

                    workSheet3.Range["A1"].AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet3.Columns.EntireColumn.AutoFit();
                    workSheet3.Range["A1", "G1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;

                    workSheet4.Range["A1"].AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet4.Columns.EntireColumn.AutoFit();
                    workSheet4.Range["A1", "H1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;
                    
                    workSheet5.Range["A1"].AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    workSheet5.Columns.EntireColumn.AutoFit();
                    workSheet5.Range["A1", "E1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;
                    
                    // Save the file, quit Excel, and display message to user.
                    // workSheet.SaveAs($@"{Environment.CurrentDirectory}\" + tableName + ".xlsx");
                    try
                    {
                        oWB.SaveAs($@"{Environment.CurrentDirectory}\" + tableName + ".xlsx");
                    }
                    catch (Exception unused)
                    {
                        excelApp.Quit();
                        return;
                    }
                    //excelApp.Quit();
                    excelApp.Visible = true;

                    System.Diagnostics.Debug.WriteLine("The tableOne.xslx file has been saved to your app folder");

                    //*********************if need assistance please ask ***************************
                    //*****************************James Alexander **************************
                }
                 
            }
            else
            {
                MessageBox.Show("No database with file name '"+user_Input.Text+"' was found. Add a new database or try again with a " +
                    "different name.", "Try again", button2, icon2);
            }

        }

        public void colorRows(Excel._Worksheet worksheet)
        {

            //counter variable
            int count = 1;

            //Iterate the rows in the used range
            foreach (Excel.Range row in worksheet.UsedRange.Rows)
            {
                //if row is even
                if (count % 2 == 0)
                {
                    row.Interior.Color = Excel.XlRgbColor.rgbWhiteSmoke;
                }
                //if row is odd
                else
                {
                    row.Interior.Color = Excel.XlRgbColor.rgbSlateGrey;
                }

                count++;
            }
        }

        private void account1TB_GotFocus(object sender, RoutedEventArgs e)
        {
            //clear any warnings, if necessary
            warningTB.Visibility = Visibility.Hidden;
            warningAN.Visibility = Visibility.Hidden;
        }
        private void user_Input_TextChanged(object sender, TextChangedEventArgs e)
        {
            passingText = user_Input.Text.Replace(" ", "");
        }




        //**************Save All table to DB*******************
        // This used to be the save button event but I turned it into a save function.
        private bool Save_DB()
        {
            LeapLogDBManager sqlTables = new LeapLogDBManager();

            
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;

            string tableName = user_Input.Text.Replace(" ", "");
            string TaccountName = tableName + "Taccount";
            string BalanceSheetName = tableName + "BalanceSheet";
            string IncomeStatementName = tableName + "IncomeStatement";
            string StatementOfOEName = tableName + "StatementOfOE";



            //**********insert into T accounts table************
            if (user_Input.Text == "")
            {
                MessageBox.Show("No table name selected.", "Error", button, icon);
                return false;
            }

            else
            {
                try
                {
                    foreach (var i in Database.TEntries)
                    {

                        sqlTables.WriteData("INSERT INTO " + TaccountName + " VALUES ('" + i.Account + "','" + i.Type + "','" + string.Join("\n", i.Debit.ToArray()) + "','" + string.Join("\n", i.Credit.ToArray()) + "','" + i.TotalDebit + "','" + i.TotalCredit + "','" + i.Balance + "')");


                    }

                    //**********insert into Balance Sheet************


                    //iterate through list of all asset accounts
                    for (int i = 0; i < Database.TEntries.Count; i++)
                    {
                        string assetAccountName = "";
                        double assetBalance = 0;
                        string loeAccountName = "";
                        double loeBalance = 0;
                        if (i < Database.BalanceData.assetsList.Count)
                        {
                            assetAccountName = Database.BalanceData.assetsList[i].Account;
                            assetBalance = Database.BalanceData.assetsList[i].Balance;
                        }
                        if (i < Database.BalanceData.loeList.Count)
                        {
                            loeAccountName = Database.BalanceData.loeList[i].Account;
                            loeBalance = Database.BalanceData.loeList[i].Balance;
                        }
                        sqlTables.WriteData("INSERT INTO " + BalanceSheetName + " VALUES ('" + assetAccountName + "','" + assetBalance + "','" + Database.BalanceData.total_assets + "','" + loeAccountName + "','" + loeBalance + "','" + Database.BalanceData.total_loe + "')");


                    }


                    //**********insert into Income Statement************

                    for (int i = 0; i < Database.TEntries.Count; i++)
                    {

                        string revenueAccountName = "";
                        double revenueBalance = 0;
                        string expenseAccountName = "";
                        double expenseBalance = 0;
                        if (i < Database.IncomeData.revenueList.Count)
                        {
                            revenueAccountName = Database.IncomeData.revenueList[i].Account;
                            revenueBalance = Database.IncomeData.revenueList[i].Balance;
                        }
                        if (i < Database.IncomeData.expenseList.Count)
                        {
                            expenseAccountName = Database.IncomeData.expenseList[i].Account;
                            expenseBalance = Database.IncomeData.expenseList[i].Balance;
                        }

                        sqlTables.WriteData("INSERT INTO " + IncomeStatementName + " VALUES ('" + revenueAccountName + "','" + revenueBalance + "','" + Database.IncomeData.total_revenue + "','" + expenseAccountName + "','" + expenseBalance + "','" + Database.IncomeData.total_expenses + "','" + Database.IncomeData.net_income + "')");
                    }

                    //**********insert into Stetemtent of Owner Equity************


                    sqlTables.WriteData("INSERT INTO " + StatementOfOEName + " VALUES ('" + Database.SoeData.start_capital + "','" + Database.SoeData.net_income + "','" + Database.SoeData.total_withdrawals + "','" + Database.SoeData.final_capital + "')");


                    //MessageBox.Show("Table data saved to database.", "Saved", button, icon);
                    return true;
                }
                catch
                {
                    MessageBox.Show("Add a database first before attempting to export.", "Error", button, icon);
                    return false;
                }
            }
        }

        private void type1CB_GotFocus(object sender, RoutedEventArgs e)
        {
            warningAT.Visibility = Visibility.Hidden;
        }

         

    }
}
