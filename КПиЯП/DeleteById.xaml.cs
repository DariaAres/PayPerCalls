using System;
using System.Windows;

namespace КПиЯП
{
    /// <summary>
    /// Логика взаимодействия для DeleteById.xaml
    /// </summary>
    public partial class DeleteById : Window
    {
        public DeleteById()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, object e)
        {
            try
            {
                MainWindow.db.Delete(Convert.ToInt32(id.Text));

                this.Visibility = Visibility.Hidden;
                this.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
