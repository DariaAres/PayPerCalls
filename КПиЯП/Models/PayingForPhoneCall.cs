using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using КПиЯП.Exceptions;
using КПиЯП.Util;

namespace КПиЯП.Models
{
    public class PayingForPhoneCall
    {
        private string lastName;
        private string phoneNumber;
        private DateTime date;
        private double rateMin;
        private string discountPercent;
        private DateTime timeIn;
        private DateTime timeOut;
        public string LastName
        {
            get { return lastName; }
            set
            {
                if (Char.IsUpper(value[0])) lastName = value;
                else throw new MyException("фамилия с маленькой буквы! ");
            }
        }

        public string PhoneNumber
        {
            get { return phoneNumber; }
            set
            {
                Regex regex = new Regex(@"\+375(44|29|25)\d{7}");
                MatchCollection matches = regex.Matches(value);
                if (matches.Count > 0 && value.Length == 13)
                {
                    foreach (Match match in matches)
                        phoneNumber = match.Value;
                }
                else throw new MyException("Неверного формата номер телефона! ");
            }
        }
        public DateTime Date
        {
            get { return date; }
            set
            {
                date = value;
            }
        }
        public double RateMin
        {
            get { return rateMin; }
            set
            {
                if (value % 100 < 100) rateMin = value;
                else throw new MyException("Тариф неверно записан");
            }
        }
        public string DiscountPercent
        {
            get { return discountPercent; }
            set
            {
                Regex regex = new Regex(@"([0-9]|([1-9][0-9])|100)%");
                MatchCollection matches = regex.Matches(value);
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                        discountPercent = match.Value.RemoveEnd(1);

                }
                else throw new MyException("Неверный формат скидки! ");
            }
        }
        public DateTime TimeIn
        {
            get { return timeIn; }
            set
            {
                timeIn = value;
            }
        }
        public DateTime TimeOut
        {
            get { return timeOut; }
            set
            {
                timeOut = value;
            }
        }
        public PayingForPhoneCall(string lastName, DateTime date, string phoneNumber, double rateMin,
            string discountPercent, DateTime timeIn, DateTime timeOut)
        {
            LastName = lastName;
            PhoneNumber = phoneNumber;
            Date = date;
            RateMin = rateMin;
            DiscountPercent = discountPercent;
            TimeIn = timeIn;
            TimeOut = timeOut;
        }
        public PayingForPhoneCall() { }

        public List<PayingForPhoneCall> Sort(List<PayingForPhoneCall> list)
        {
            var sortedList = from p in list
                             orderby p.LastName, p.Date
                             select p;

            return sortedList.ToList();
        }

        

        public int SearchPhones(List<PayingForPhoneCall> list)
        {
            int i = 0;
            foreach (var p in list)
            {
                if (p.PhoneNumber == PhoneNumber) i++;
            }
            return i;
        }

        public string SortDescendin(List<PayingForPhoneCall> list)
        {
            string s = "";
            var orderedNumbers = list.OrderByDescending(n => n.TimeOut.Subtract(n.TimeIn));
            foreach (var i in orderedNumbers)
                s+=i;
            return s;
        }

        public double Price()
        {
            return RateMin - (RateMin / 100 * Convert.ToInt32(DiscountPercent));
        }
    }
}
