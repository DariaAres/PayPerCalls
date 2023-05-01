using System;
using System.Windows;

namespace КПиЯП
{
    /// <summary>
    /// Логика взаимодействия для DeleteInRange.xaml
    /// </summary>
    public partial class DeleteInRange : Window
    {
        public DeleteInRange()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.db.DeleteInRange(Convert.ToInt32(first.Text), Convert.ToInt32(last.Text));

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
