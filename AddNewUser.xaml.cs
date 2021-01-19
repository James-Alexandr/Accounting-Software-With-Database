using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace LeapLog
{
    /// <summary>
    /// Interaction logic for AddNewUser.xaml
    /// </summary>
    public partial class AddNewUser : Window
    {
       
        public AddNewUser()
        {
            InitializeComponent();
        }



        private void exitButton_Click_1(object sender, RoutedEventArgs e)
        {
            // this.Close();
            //Process.GetCurrentProcess().Close();
            Environment.Exit(Environment.ExitCode);




        }

        private void backToLogin_Click(object sender, RoutedEventArgs e)
        {
            LoginScreen MS = new LoginScreen();
            MS.Show();
            this.Hide();
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                //our code here
                // this.Close();
                Environment.Exit(Environment.ExitCode);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            List<String> list = new List<String>();

            LoginManager sqlTables = new LoginManager();

            string messageBoxText = "Username and password fields cannot be null or empty.";
            string messageBoxText2 = "Username and password have been added.";
            string caption = "Wrong Input";
            string caption2 = "User Added";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Information;

            // sqlTables.ReadData("Select * from AdminLogin ");

            SqlConnection conn = new SqlConnection($@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={Environment.CurrentDirectory}\LOGINDB.MDF;Integrated Security=True");
            string query = "Select * from UserLogin ";
            SqlDataAdapter sda = new SqlDataAdapter(query, conn);
            DataTable dataTable = new DataTable();
            sda.Fill(dataTable);

            string name;

            foreach (DataRow row in dataTable.Rows)
            {
                name = row["username"].ToString().Trim();
                System.Diagnostics.Debug.WriteLine(name);
                list.Add(name);
            }


            // if (name.Equals(newUsername.Text.Trim(), StringComparison.CurrentCultureIgnoreCase))

            //<<-----------this code compares user input string with strings in database------------>>
            if (list.FindIndex(x => x.Equals(newUsername.Text.Trim(), StringComparison.CurrentCultureIgnoreCase)) != -1)
            
                    MessageBox.Show("Username already taken", "Try again", button, icon);
                     

                //<<---------this code adds the string username and password to the db
            else
            {
                string username = newUsername.Text.Replace(" ", "");
                string password = newPass.Password.ToString();


                if (String.IsNullOrEmpty(username) || username == "")
                {
                    MessageBox.Show(messageBoxText, caption, button, icon);
                }
                else if (String.IsNullOrEmpty(password) || password == "")
                {
                    MessageBox.Show(messageBoxText, caption, button, icon);
                }
                else
                {
                    sqlTables.WriteData(" INSERT INTO UserLogin VALUES ('" + username + "','" + password + "')  ");
                    MessageBox.Show(messageBoxText2, caption2, button, icon);
                }

            }

            //<<--------this clears the textboxes after inserting input----------->>
            //string l = "";
            newUsername.Clear();
            //labelPrint.Content = l.ToString();
            newPass.Clear();
            newUsername.Focus();

        }
        //enter the program by pressing enter when the tab is on the Add Button
        private void createButton_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.Key == Key.Enter)
            {
                createButton_Click(sender, e);

            }
        }

        //enter the program by pressing enter when the tab is on the Passwordbox
        private void newPass_KeyUp(object sender, KeyEventArgs e)
        {
            createButton_KeyUp(sender, e);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
    }
}
