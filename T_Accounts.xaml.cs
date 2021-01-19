using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;

namespace LeapLog
{
     /// <summary>
     /// Go ahead and write a summary for T_Accounts.xaml here!
     /// </summary>
     public partial class T_Accounts : UserControl
     {
        public static string destinationText = Journal.passingText;

        public T_Accounts()
          {
               InitializeComponent();
          }

        //method that refreshes the t-account grid so that it is updated with current journal entries
        public void Refresh()
        {
            //clear grid
            a_Grid.Items.Clear();
            l_Grid.Items.Clear();
            oe_Grid.Items.Clear();

            //add t-accounts to respective grids
            for (int i = 0; i < Database.TEntries.Count; i++)
            {
                if (Database.TEntries[i].Type.Equals("Asset"))
                {
                    a_Grid.Items.Add(Database.TEntries[i]);
                }
                else if (Database.TEntries[i].Type.Equals("Liability"))
                {
                    l_Grid.Items.Add(Database.TEntries[i]);
                }
                else if (Database.TEntries[i].Type.Equals("Owners Equity") ||
                        Database.TEntries[i].Type.Equals("Revenue") ||
                        Database.TEntries[i].Type.Equals("Expense") ||
                        Database.TEntries[i].Type.Equals("Withdrawal"))
                {
                    oe_Grid.Items.Add(Database.TEntries[i]);
                }
            }
        }

        //method that processes temporary t-accounts by either
        //adding them to the t-account list or updating t-account list
        //with the temporary t-account data
        public static void add_taccs(List<Entry_tacc> tempAccounts)
        {
            //process each temporary account
            Entry_tacc acc1 = tempAccounts[0];
            Entry_tacc acc2 = tempAccounts[1];

            //if database is empty of t-accounts
            if (Database.TEntries.Count == 0)
            {
                //add accounts to database
                Database.TEntries.Add(acc1);
                Database.TEntries.Add(acc2);
            }

            //if database is not empty,
            //compare and either update an existing account
            //or add the account to the database
            else
            {
                //variables to keep track of whether or not a t-account has been updated
                bool updated1 = false;
                bool updated2 = false;

                for (int i = 0; i < Database.TEntries.Count; i++)
                {
                    //if first t-account exists already
                    if (acc1.Account.Equals(Database.TEntries[i].Account))
                    {
                        //add the temporary t-account debit and credit to the original
                        Database.TEntries[i].Debit.Add(acc1.Debit[0]);
                        Database.TEntries[i].Debit.Add(acc1.Credit[0]);

                        //get sum of all current debits
                        double sumDebit = 0;
                        for (int j = 0; j < Database.TEntries[i].Debit.Count; j++)
                        {
                            sumDebit += Database.TEntries[i].Debit[j];
                        }

                        //get sum of all current credits
                        double sumCredit = 0;
                        for (int j = 0; j < Database.TEntries[i].Credit.Count; j++)
                        {
                            sumCredit += Database.TEntries[i].Credit[j];
                        }

                        //if it is an asset, withdrawal, or expense type
                        if (Database.TEntries[i].Type.Equals("Asset") || Database.TEntries[i].Type.Equals("Expense") || Database.TEntries[i].Type.Equals("Withdrawal"))
                        {
                            //update the original account's balance by subtracting the credit from the debit
                            Database.TEntries[i].Balance = sumDebit - sumCredit;
                        }
                        //if it is a liability, revenue, or owner's equity type
                        else
                        {
                            //update the original account's balance by subtracting the debit from the credit
                            Database.TEntries[i].Balance = sumCredit - sumDebit;
                        }

                        Database.TEntries[i].TotalCredit = sumCredit;
                        Database.TEntries[i].TotalDebit = sumDebit;

                        updated1 = true;
                    }
                    //if end of list reached and the t-account does not exist,
                    //update the account's properties and add account to database
                    else if (i == Database.TEntries.Count - 1 && !(acc1.Account.Equals(Database.TEntries[i].Account)) && !updated1)
                    {
                        Database.TEntries.Add(acc1);
                        break;
                    }

                }
                for (int i = 0; i < Database.TEntries.Count; i++)
                {
                    //if first t-account exists already
                    if (acc2.Account.Equals(Database.TEntries[i].Account))
                    {
                        //add the temporary t-account credit and debit to the original
                        Database.TEntries[i].Debit.Add(acc2.Debit[0]);
                        Database.TEntries[i].Credit.Add(acc2.Credit[0]);

                        //get sum of all current debits
                        double sumDebit = 0;
                        for (int j = 0; j < Database.TEntries[i].Debit.Count; j++)
                        {
                            sumDebit += Database.TEntries[i].Debit[j];
                        }

                        //get sum of all current credits
                        double sumCredit = 0;
                        for (int j = 0; j < Database.TEntries[i].Credit.Count; j++)
                        {
                            sumCredit += Database.TEntries[i].Credit[j];
                        }

                        //if it is an asset, withdrawal, or expense type
                        if (Database.TEntries[i].Type.Equals("Asset") || Database.TEntries[i].Type.Equals("Expense") || Database.TEntries[i].Type.Equals("Withdrawal"))
                        {
                            //update the original account's balance by subtracting the credit from the debit
                            Database.TEntries[i].Balance = sumDebit - sumCredit;
                        }
                        //if it is a liability, revenue, or owner's equity type
                        else
                        {
                            //update the original account's balance by subtracting the debit from the credit
                            Database.TEntries[i].Balance = sumCredit - sumDebit;
                        }

                        Database.TEntries[i].TotalCredit = sumCredit;
                        Database.TEntries[i].TotalDebit = sumDebit;

                        updated2 = true;
                    }
                    //else, if end of list reached and the t-account does not exist,
                    //add account to list
                    else if (i == Database.TEntries.Count - 1 && !(acc2.Account.Equals(Database.TEntries[i].Account)) && !updated2)
                    {
                        Database.TEntries.Add(acc2);
                        break;
                    }
                }
            }

            //once list of accounts is updated, update list of accounts for balance sheet and income statement

            //clear existing data
            Database.BalanceData.assetsList.Clear();
            Database.BalanceData.loeList.Clear();
            Database.IncomeData.revenueList.Clear();
            Database.IncomeData.expenseList.Clear();

            //remove zeros from list of debits and credits
            for (int i = 0; i < Database.TEntries.Count; i++)
            {
                for (int j = 0; j < Database.TEntries[i].Debit.Count; j++)
                {
                    if (Database.TEntries[i].Debit[j] == 0)
                    {
                        Database.TEntries[i].Debit.RemoveAt(j);
                    }
                }
                for (int j = 0; j < Database.TEntries[i].Credit.Count; j++)
                {
                    if (Database.TEntries[i].Credit[j] == 0)
                    {
                        Database.TEntries[i].Credit.RemoveAt(j);
                    }
                }
            }

                //enter updated data
                for (int i = 0; i < Database.TEntries.Count; i++)
            {
                if (Database.TEntries[i].Type == "Asset")
                {
                    Database.BalanceData.assetsList.Add(Database.TEntries[i]);
                }
                else
                {
                    Database.BalanceData.loeList.Add(Database.TEntries[i]);
                }

                if (Database.TEntries[i].Type == "Revenue")
                {
                    Database.IncomeData.revenueList.Add(Database.TEntries[i]);
                }
                else if (Database.TEntries[i].Type == "Expense")
                {
                    Database.IncomeData.expenseList.Add(Database.TEntries[i]);
                }
            }
        }

        //method that creates two temporary t-accounts that corresponds to a journal entry
        public static List<Entry_tacc> get_taccs(Entry entry)
        {
            //create a new list of t-account entries
            List<Entry_tacc> tempAccounts = new List<Entry_tacc>();

            //create two t-accounts and populate with data from journal entry

            //account 1
            Entry_tacc acc1 = new Entry_tacc();
            acc1.Date = entry.Date;
            acc1.Account = entry.Account1.Trim();
            acc1.Type = entry.Type1;
            acc1.Debit.Add(entry.Debit);
            acc1.Credit.Add(0);
            acc1.TotalDebit = entry.Debit;

            //calculate balance
            if (acc1.Type == "Asset" || acc1.Type == "Expense" || acc1.Type == "Withdrawal")
            {
                acc1.Balance = acc1.Debit[0] - acc1.Credit[0];
            }
            else
            {
                acc1.Balance = acc1.Credit[0] - acc1.Debit[0];
            }

            //account 2
            Entry_tacc acc2 = new Entry_tacc();
            acc2.Date = entry.Date;
            acc2.Account = entry.Account2.Trim();
            acc2.Type = entry.Type2;
            acc2.Debit.Add(0);
            acc2.Credit.Add(entry.Credit);
            acc2.TotalCredit = entry.Credit;

            //calculate balance
            if (acc2.Type == "Asset" || acc2.Type == "Expense" || acc2.Type == "Withdrawal")
            {
                acc2.Balance = acc2.Debit[0] - acc2.Credit[0];
            }
            else
            {
                acc2.Balance = acc2.Credit[0] - acc2.Debit[0];
            }

            //add t-accounts to list
            tempAccounts.Add(acc1);
            tempAccounts.Add(acc2);

            //return the t-accounts
            return tempAccounts;
          }

        //method that occurs when a cell in any of the grids is clicked: it
        //gets the cell's respective account's name, data, and balance
        private void DataGridCell_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //open a new window for the account
            SingleAccount accountWindow = new SingleAccount();
            accountWindow.Show();

            //get the account name and set in new window
            DataGridCell cell = (DataGridCell)sender;
            string cellName = cell.ToString();
            int index = cellName.IndexOf(":");
            string accountName = cellName.Substring(index + 2);
            accountWindow.accountName.Content = accountName;

            //get account data
            Entry_tacc tacc = new Entry_tacc();

            for (int i = 0; i < Database.TEntries.Count; i++)
            {
                if (accountName.Equals(Database.TEntries[i].Account))
                {
                    tacc.Debit.AddRange(Database.TEntries[i].Debit);
                    tacc.Credit.AddRange(Database.TEntries[i].Credit);
                    tacc.Balance = Database.TEntries[i].Balance;
                }
            }

            //set account data in new window
            
            for (int i = 0; i < tacc.Debit.Count; i++)
            {
                Account account = new Account();
                account.Debit = tacc.Debit[i];
                if (account.Debit != 0)
                {
                    accountWindow.debitGrid.Items.Add(account);
                }
                
            }
            for (int i = 0; i < tacc.Credit.Count; i++)
            {
                Account account = new Account();
                account.Credit = tacc.Credit[i];
                if (account.Credit != 0)
                {
                    accountWindow.creditGrid.Items.Add(account);
                }
            }

            accountWindow.balanceLbl.Content = tacc.Balance.ToString();

            //to clear duplicate cells that pop up after user clicks a cell
            Refresh();
        }

        //method for when help feature button is clicked
        private void taccHelpButton_Click(object sender, RoutedEventArgs e)
        {
            if (taccHelpWindow.Visibility == Visibility.Collapsed)
            {
                taccHelpWindow.Visibility = Visibility.Visible;
            }
            else
            {
                taccHelpWindow.Visibility = Visibility.Collapsed;
            }
        }
    }
}
