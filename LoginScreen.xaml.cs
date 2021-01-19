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
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Data;
using System.IO;

namespace LeapLog
{
    /// <summary>
    /// this is the git loginScreen Branch
    /// Interaction logic for LoginScreen.xaml
    /// </summary>
    ///
    //color #FFB4E5F0  and #FFB4DDF0
    public partial class LoginScreen : Window
    {

     

        public object MessageBoxIcon { get; private set; }

        public LoginScreen()
        {
            InitializeComponent();
        }

        /*Button Click login*/

         private void button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;

            //database funtionality for user login credentials

            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory); //or set executing Assembly location path in param

           
            UserName.Focus();
            SqlConnection conn = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\LOGINDB.MDF;Integrated Security=True");
            string query = "Select * from UserLogin where username = '" + UserName.Text.Trim() + "' and password = '" + Pass.Password.ToString() + "'";
            SqlDataAdapter sda = new SqlDataAdapter(query, conn);
            DataTable dataTable = new DataTable();
            sda.Fill(dataTable);
            if (dataTable.Rows.Count == 1)
            {
                MainWindow MS = new MainWindow();
                MS.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show(this, "The username or password is incorrect.", "Error", button, icon);//, MessageBoxButtons.OK,
                                                                                    // MessageBoxIcon.Information);
                UserName.Clear();
                Pass.Clear();
                UserName.Focus();
            }




            /* CODE USED FOR TESTING
             * if (UserName.Text == "leaplog" && Pass.Password == "leaplog")
             {
                 MainWindow MS = new MainWindow();
                 MS.Show();
                 this.Hide();
             }
             else
             {
                 MessageBox.Show(this, "Wrong User Name or Password", "Information");//, MessageBoxButtons.OK,
                // MessageBoxIcon.Information);
                 UserName.Clear();
                 Pass.Clear();
                 UserName.Focus();
             }*/
        }

        /*Enter button login using tab as selector*/
        private void enterButton_KeyUp(object sender, KeyEventArgs e)
        {


              if (e.Key == Key.Enter)
              {
                button_Click(sender, e);

              }



        }

        //enter button functionality for pass box
        private void Pass_KeyUp_1(object sender, KeyEventArgs e)
        {
            enterButton_KeyUp(sender, e);


        }


        //clear the box when mouse clicks
        private void UserName_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            UserName.Clear();
        }

        //clear the box when mouse clicks
        private void Pass_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Pass.Clear();
        }


        //clear textbox when tab is pressed
        private void UserName_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab)
            {
                //my code here
                UserName.Clear();

            }
        }
        //clear textbox when tab is pressed
        private void Pass_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab)
            {
                //my code here
                Pass.Clear();

            }
        }

        //exit button
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //this.Close();
            Environment.Exit(Environment.ExitCode);

        }


        //close the window when esc is pressed
        private void Window_KeyUp_1(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                //our code here
                // this.Close();
                Environment.Exit(Environment.ExitCode);

            }
        }
        //sends user to the admin login
        private void newUser_Click(object sender, RoutedEventArgs e)
        {
            AdminLogin MS = new AdminLogin();
            MS.Show();
            this.Hide();
        }

        //exits application after button x closed
        private void loginWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
             System.Windows.Application.Current.Shutdown();
        }

        private void loginWindow_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
    }
}
