using System;

namespace КПиЯП.Util
{
    public static class StringExtensions
    {
        public static String RemoveEnd(this String str, int len)
        {
            if (str.Length < len)
            {
                return string.Empty;
            }

            return str.Remove(str.Length - len);
        }

        public static string ReplaceAt(this string input, int index, char newChar)
        {
            if (input == null)
            {
                throw new ArgumentNullException("input");
            }
            char[] chars = input.ToCharArray();
            chars[index] = newChar;
            return new string(chars);
        }
    }
}
