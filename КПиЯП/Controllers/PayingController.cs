using System;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows;
using КПиЯП.Models;
using КПиЯП.Util;

namespace КПиЯП.Controllers
{
    public class PayingController
    {
        static string connectionString = "Server=DESKTOP-55DDMMD\\SQLEXPRESS;Database=payings2bd;Trusted_Connection=True;TrustServerCertificate=True;";

        public async Task Insert(PayingForPhoneCall paying)
        {
            string sqlExpression = "insert into Payings(LastName, Phone, [Date], Rate, Discount, [TimeIn], [TimeOut], CallLength) values (@lastName, @phone, @date, @rate, @discount, @timeIn, @timeOut, @length);";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    command.Parameters.AddWithValue("@lastName", paying.LastName);
                    command.Parameters.AddWithValue("@phone", paying.PhoneNumber);
                    command.Parameters.AddWithValue("@date", paying.Date.Date);
                    command.Parameters.AddWithValue("@rate", paying.RateMin.ToString().ReplaceAt(2, '.'));
                    command.Parameters.AddWithValue("@discount", paying.DiscountPercent);
                    command.Parameters.AddWithValue("@timeIn", paying.TimeIn);
                    command.Parameters.AddWithValue("@timeOut", paying.TimeOut);
                    command.Parameters.AddWithValue("@length", paying.TimeOut.Subtract(paying.TimeIn));

                    int number = command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Insert error");
            }
        }

        //public async Task Update()
        //{
        //    //user.ToUpdate();
        //    string sqlExpression = "UPDATE _ SET password='" + "'";
        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();
        //        SqlCommand command = new SqlCommand(sqlExpression, connection);
        //        int number = command.ExecuteNonQuery();
        //        Console.WriteLine("Обновлено объектов: {0}", number);
        //    }
        //}

        public async Task DeleteInRange(int first, int last)
        {
            int count = 0;
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    for (int i = first; i <= last; i++)
                    {
                        string sqlExpression = "use payingsbd; DELETE  FROM Payings WHERE id=" + i;
                        count++;
                        
                        SqlCommand command = new SqlCommand(sqlExpression, connection);
                        int number = command.ExecuteNonQuery();
                        count += number;
                    }
                }
                MessageBox.Show("Удалено объектов: " + count);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete in range error");
            } 
        }
        public async Task Delete(int id)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string sqlExpression = "use payingsbd; DELETE  FROM Payings WHERE id=" + id;

                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    int number = command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete in range error");
            }
        }
    }
}
