using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LeapLog
{
    /// <summary>
    /// The main window represents is used to navigate between User Controls / program sections.
    /// </summary>
    /// this is the testing new feature clone

    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        //adding shutdown funtionality when closing main window
        private void Window_Closing_1(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        //when switching tabs, t-account section refreshes
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            t_account_control.Refresh();
            balance_sheet_control.Refresh();
            income_statement.Refresh();
            statementofownersequity.Refresh();
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (gettingStartedWindow.Visibility == Visibility.Collapsed)
            {
                gettingStartedWindow.Visibility = Visibility.Visible;
            }
            else {
                gettingStartedWindow.Visibility = Visibility.Collapsed;
            }
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.Close();
        }

    }

}
