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

namespace LeapLog
{
    /// <summary>
    /// Interaction logic for Dashboard.xaml
    /// </summary>
    public partial class Dashboard : UserControl
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        //switches explanation views
        private void accBtn_Click(object sender, RoutedEventArgs e)
        {
            accGrid.Visibility = Visibility.Visible;
            assetGrid.Visibility = Visibility.Collapsed;
            liabilityGrid.Visibility = Visibility.Collapsed;
            oeGrid.Visibility = Visibility.Collapsed;
        }

        private void assetBtn_Click(object sender, RoutedEventArgs e)
        {
            accGrid.Visibility = Visibility.Collapsed;
            assetGrid.Visibility = Visibility.Visible;
            liabilityGrid.Visibility = Visibility.Collapsed;
            oeGrid.Visibility = Visibility.Collapsed;
        }

        private void liabilityBtn_Click(object sender, RoutedEventArgs e)
        {
            accGrid.Visibility = Visibility.Collapsed;
            assetGrid.Visibility = Visibility.Collapsed;
            liabilityGrid.Visibility = Visibility.Visible;
            oeGrid.Visibility = Visibility.Collapsed;
        }

        private void oeBtn_Click(object sender, RoutedEventArgs e)
        {
            accGrid.Visibility = Visibility.Collapsed;
            assetGrid.Visibility = Visibility.Collapsed;
            liabilityGrid.Visibility = Visibility.Collapsed;
            oeGrid.Visibility = Visibility.Visible;
        }
    }
}
