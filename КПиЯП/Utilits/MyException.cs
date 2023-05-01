using System;
using System.Windows;

namespace КПиЯП.Exceptions
{
    public class MyException : Exception
    {
        public MyException(string s) : base(s) { MessageBox.Show(s, "Error"); }
    }
}
