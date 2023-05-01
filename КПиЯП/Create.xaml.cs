using System;
using System.Windows;
using КПиЯП.Models;

namespace КПиЯП
{
    /// <summary>
    /// Логика взаимодействия для Create.xaml
    /// </summary>
    public partial class Create : Window
    {
        public Create()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var paying = new PayingForPhoneCall();

                paying.LastName = lastname.Text;
                paying.Date = Convert.ToDateTime(date.Text);
                paying.PhoneNumber = phone.Text;
                paying.RateMin = Convert.ToDouble(rake.Text);
                paying.DiscountPercent = discount.Text + "%";
                paying.TimeIn = Convert.ToDateTime(timein.Text);
                paying.TimeOut = Convert.ToDateTime(timeout.Text);

                MainWindow.db.Insert(paying);

                this.Visibility = Visibility.Hidden;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
