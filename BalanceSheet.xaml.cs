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
	 /// Interaction logic for BalanceSheet.xaml
	 /// </summary>
	 public partial class BalanceSheet : UserControl
	 {
		public static string destinationText = Journal.passingText;
		
		public BalanceSheet()
		  {
			   InitializeComponent();
			   Database.select_date = DateTime.Now;
				DateTime now = DateTime.Now;
				string format = "MMMM dd, yyyy";
				selectDateBox.Text = now.ToString(format);
		}

		// method that refreshes the balancesheet grid so that it is updated with current t-account entries
		public void Refresh()
		{
			// Clear all grids
			entryGridA.Items.Clear();
			entryGridTA.Items.Clear();
			entryGridLO.Items.Clear();
			entryGridTLO.Items.Clear();

			double Total_assests = 0;
			double Total_LOE = 0;

			//add t-accounts to respective grids
			for (int i = 0; i < Database.TEntries.Count; i++)
			{
				Entry_tacc _tacc = Database.TEntries[i];
				DateTime td = _tacc.Date;
				DateTime date = Database.select_date;
				if (td.Day == date.Day && td.Month == date.Month && td.Year == date.Year)
				{
					string type = _tacc.Type;
					switch (type)
					{
						case "Asset":
							entryGridA.Items.Add(_tacc);
							Total_assests += _tacc.Balance;
							break;
						case "Liability":
							entryGridLO.Items.Add(_tacc);
							Total_LOE += _tacc.Balance;
							break;
						case "Expense":
						case "Withdrawal":
							entryGridLO.Items.Add(_tacc);
							Total_LOE -= _tacc.Balance;
							break;
						case "Revenue":
						case "Owners Equity":
							entryGridLO.Items.Add(_tacc);
							Total_LOE += _tacc.Balance;
							break;
					}
				}
			}

			// Add the totals
			entryGridTA.Items.Add(new Total() { total = Total_assests });
			entryGridTLO.Items.Add(new Total() { total = Total_LOE });

			//add balance sheet data into database
			Database.BalanceData.total_assets = Total_assests;
			Database.BalanceData.total_loe = Total_LOE;
		}

		//method for when help feature button is clicked
		private void balanceHelpButton_Click(object sender, RoutedEventArgs e)
		{
			if (balanceHelpWindow.Visibility == Visibility.Collapsed)
			{
				balanceHelpWindow.Visibility = Visibility.Visible;
			}
			else
			{
				balanceHelpWindow.Visibility = Visibility.Collapsed;
			}
		}

		private void selectDateBox_SelectionChanged(object sender, RoutedEventArgs e)
		{
			DateTime.TryParse(selectDateBox.Text, out Database.select_date);
			Refresh();
		}
	}
	}

