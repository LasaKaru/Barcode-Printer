using System.Windows;

namespace BarcodePrinterApp
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            if (txtUsername.Text == "ranjith" && txtPassword.Password == "ranjith")
            {
                // Correct credentials - open main app and close login
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                this.Close();
            }
            else
            {
                // Wrong credentials - show error
                txtError.Visibility = Visibility.Visible;
            }
        }
    }
}