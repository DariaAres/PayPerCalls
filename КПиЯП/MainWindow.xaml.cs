using System;
using System.Linq;
using System.Windows;
using КПиЯП.Controllers;
using КПиЯП.Models;
using КПиЯП.Models;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using FakeItEasy;
using System.Collections;

namespace КПиЯП
{
    public partial class MainWindow : System.Windows.Window
    {
        public static PayingController db = new PayingController();
        ArrayList queries = new ArrayList();

        payings2bdEntities1 dataEntities = new payings2bdEntities1();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Add10();
            var query =
            from p in dataEntities.Payings
                select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut };
            dataGrid1.ItemsSource = query.ToList();
            var qe = query.ToList();
        }

        private void Add10()
        {
            db.Insert(new PayingForPhoneCall("Koul", new DateTime(2015, 10, 20, 18, 30, 0), "+375293645890", 43.23, "2%", new DateTime(2015, 7, 20, 18, 30, 0), new DateTime(2015, 7, 20, 18, 35, 0)));
            db.Insert(new PayingForPhoneCall("Alec", new DateTime(2013, 7, 10, 18, 30, 0), "+375293645890", 43.23, "2%", new DateTime(2015, 7, 20, 18, 30, 0), new DateTime(2015, 7, 20, 18, 35, 0)));
            db.Insert(new PayingForPhoneCall("Kapoor", new DateTime(2019, 12, 12, 18, 30, 0), "+375293645890", 43.23, "2%", new DateTime(2015, 7, 20, 18, 30, 0), new DateTime(2015, 7, 20, 18, 35, 0)));
            db.Insert(new PayingForPhoneCall("Bahar", new DateTime(2011, 7, 19, 18, 30, 0), "+375293645890", 43.23, "9%", new DateTime(2015, 7, 20, 18, 30, 0), new DateTime(2015, 7, 20, 18, 35, 0)));
            db.Insert(new PayingForPhoneCall("Jonos", new DateTime(2015, 10, 20, 18, 30, 0), "+375447623498", 37.34, "9%", new DateTime(2015, 7, 20, 20, 10, 0), new DateTime(2015, 7, 20, 20, 20, 0)));
            db.Insert(new PayingForPhoneCall("Mars", new DateTime(2013, 7, 10, 18, 30, 0), "+375253798250", 20.23, "9%", new DateTime(2015, 7, 20, 18, 31, 0), new DateTime(2015, 7, 20, 18, 31, 10)));
            db.Insert(new PayingForPhoneCall("King", new DateTime(2019, 12, 12, 18, 30, 0), "+375293473458", 43.83, "89%", new DateTime(2015, 7, 20, 17, 30, 0), new DateTime(2015, 7, 20, 17, 35, 0)));
            db.Insert(new PayingForPhoneCall("Kristy", new DateTime(2010, 9, 14, 18, 30, 0), "+375447623498", 37.34, "89%", new DateTime(2015, 7, 20, 20, 10, 0), new DateTime(2015, 7, 20, 20, 20, 0)));
            db.Insert(new PayingForPhoneCall("Air", new DateTime(2019, 12, 12, 18, 30, 0), "+375253798250", 20.23, "9%", new DateTime(2015, 7, 20, 18, 31, 0), new DateTime(2015, 7, 20, 18, 31, 10)));
            db.Insert(new PayingForPhoneCall("Hadid", new DateTime(2015, 10, 20, 18, 30, 0), "+375293695490", 98.23, "2%", new DateTime(2015, 7, 20, 2, 32, 0), new DateTime(2015, 7, 20, 3, 35, 0)));
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Create w = new Create();
            w.Show();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            DeleteById w = new DeleteById();
            w.Show();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            DeleteInRange w = new DeleteInRange();
            w.Show();
        }
        
        private void OrderByNameAndDate(object sender, RoutedEventArgs e)
        {
            var query =
            from p in dataEntities.Payings
            orderby p.LastName
            orderby p.Date
            select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut };
            dataGrid1.ItemsSource = query.ToList();
           
        }

        private void FindNotUniqPhones(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.Phone).Where(i => i.Count() != 1);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "Phone";
        }

        private void OrderBySubstractingOfTimes(object sender, RoutedEventArgs e)
        {
            var query =
            from p in dataEntities.Payings
            orderby p.CallLength
            select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut, p.CallLength};
            dataGrid1.ItemsSource = query.ToList();
        }

        private void HigherPriceWhithDiscount(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.OrderByDescending(i => (i.Rate-(i.Rate/100*i.Discount))).FirstOrDefault();
            MessageBox.Show(query.Rate.ToString(), "Higher rate whith discount");
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.LastName);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "LastName";
        }
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.Phone);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "Phone";
        }
        
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.Date);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "Date";
        }
        
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.Rate);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "Rate";
        }
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.Discount);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "Discount";
        }

        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.TimeIn);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "TimeIn";
        }
        
        private void Button_Click_10(object sender, RoutedEventArgs e)
        {
            var query = dataEntities.Payings.GroupBy(u => u.TimeOut);
            dataGrid1.ItemsSource = query.ToList();
            dataGrid1.Columns[0].Header = "TimeOut";
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            Document document = application.Documents.Add();

            Microsoft.Office.Interop.Word.Paragraph payingParagrap = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRange = payingParagrap.Range;
            payingRange.Text = "Таблицы оплат за телефонные звонки";
            payingParagrap.set_Style("Обычный");
            payingRange.InsertParagraphAfter();
            var query =
            from p in dataEntities.Payings
            select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut };
            var allPayings = query.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrap = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRange = tableParagrap.Range;
            Microsoft.Office.Interop.Word.Table priceTable = document.Tables.Add(tableRange, allPayings.Count() + 1, 7);
            priceTable.Borders.InsideLineStyle = priceTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            priceTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            priceTable.Borders.InsideLineStyle = priceTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            priceTable.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRange;
            cellRange = priceTable.Cell(1, 1).Range;
            cellRange.Text = "Фамилия";
            cellRange = priceTable.Cell(1, 2).Range;
            cellRange.Text = "Дата оплаты";
            cellRange = priceTable.Cell(1, 3).Range;
            cellRange.Text = "Телефон";
            cellRange = priceTable.Cell(1, 4).Range;
            cellRange.Text = "Тариф";
            cellRange = priceTable.Cell(1, 5).Range;
            cellRange.Text = "Скидка";
            cellRange = priceTable.Cell(1, 6).Range;
            cellRange.Text = "Время начала разговора";
            cellRange = priceTable.Cell(1, 7).Range;
            cellRange.Text = "Время конца разговора";
            for (int i = 0; i < allPayings.Count(); ++i)
            {
                var currentphone = allPayings.ElementAt(i);
                cellRange = priceTable.Cell(i + 2, 1).Range;
                cellRange.Text = currentphone.LastName;
                cellRange = priceTable.Cell(i + 2, 2).Range;
                cellRange.Text = currentphone.Date.ToString();
                cellRange = priceTable.Cell(i + 2, 3).Range;
                cellRange.Text = currentphone.Phone;
                cellRange = priceTable.Cell(i + 2, 4).Range;
                cellRange.Text = currentphone.Rate.ToString();
                cellRange = priceTable.Cell(i + 2, 5).Range;
                cellRange.Text = currentphone.Discount.ToString() + "%";
                cellRange = priceTable.Cell(i + 2, 6).Range;
                cellRange.Text = currentphone.TimeIn.ToString();
                cellRange = priceTable.Cell(i + 2, 7).Range;
                cellRange.Text = currentphone.TimeOut.ToString();
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapOrdByND = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeOrdByNA = payingParagrapOrdByND.Range;
            payingRangeOrdByNA.Text = "Сортировка по фамилии и дате";
            payingParagrapOrdByND.set_Style("Обычный");
            payingRangeOrdByNA.InsertParagraphAfter();

            var queryOrdByND =
            from p in dataEntities.Payings
            orderby p.LastName
            orderby p.Date
            select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut };
            var OrdByND = queryOrdByND.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapOrdByND = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeOrdByND = tableParagrapOrdByND.Range;
            Microsoft.Office.Interop.Word.Table payingTableOrdByND = document.Tables.Add(tableRangeOrdByND, OrdByND.Count() + 1, 7);
            payingTableOrdByND.Borders.InsideLineStyle = payingTableOrdByND.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableOrdByND.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableOrdByND.Borders.InsideLineStyle = payingTableOrdByND.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableOrdByND.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeOrdByND;
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 1).Range;
            cellRangeOrdByND.Text = "Фамилия";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 2).Range;
            cellRangeOrdByND.Text = "Дата оплаты";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 3).Range;
            cellRangeOrdByND.Text = "Телефон";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 4).Range;
            cellRangeOrdByND.Text = "Тариф";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 5).Range;
            cellRangeOrdByND.Text = "Скидка";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 6).Range;
            cellRangeOrdByND.Text = "Время начала разговора";
            cellRangeOrdByND = payingTableOrdByND.Cell(1, 7).Range;
            cellRangeOrdByND.Text = "Время конца разговора";
            for (int i = 0; i < OrdByND.Count(); ++i)
            {
                var currentphone = OrdByND.ElementAt(i);
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 1).Range;
                cellRangeOrdByND.Text = currentphone.LastName;
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 2).Range;
                cellRangeOrdByND.Text = currentphone.Date.ToString();
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 3).Range;
                cellRangeOrdByND.Text = currentphone.Phone;
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 4).Range;
                cellRangeOrdByND.Text = currentphone.Rate.ToString();
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 5).Range;
                cellRangeOrdByND.Text = currentphone.Discount.ToString() + "%";
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 6).Range;
                cellRangeOrdByND.Text = currentphone.TimeIn.ToString();
                cellRangeOrdByND = payingTableOrdByND.Cell(i + 2, 7).Range;
                cellRangeOrdByND.Text = currentphone.TimeOut.ToString();
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapNUP = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeNUP = payingParagrapNUP.Range;
            payingRangeNUP.Text = "Повторяющиеся телефоны";
            payingParagrapNUP.set_Style("Обычный");
            payingRangeNUP.InsertParagraphAfter();
            var queryNUP = dataEntities.Payings.GroupBy(u => u.Phone).Where(i => i.Count() != 1);
            var NUP = queryNUP.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapNUP = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeNUP = tableParagrapNUP.Range;
            Microsoft.Office.Interop.Word.Table payingTableNUP = document.Tables.Add(tableRangeNUP, NUP.Count() + 1, 1);
            payingTableNUP.Borders.InsideLineStyle = payingTableNUP.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableNUP.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableNUP.Borders.InsideLineStyle = payingTableNUP.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableNUP.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeNUP;
            cellRangeNUP = payingTableNUP.Cell(1, 1).Range;
            cellRangeNUP.Text = "Телефон";
            for (int i = 0; i < NUP.Count(); ++i)
            {
                cellRangeNUP = payingTableNUP.Cell(i + 2, 1).Range;
                cellRangeNUP.Text = NUP.ElementAt(i).Key;
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapOBL = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeOBL = payingParagrapOBL.Range;
            payingRangeOBL.Text = "Сортировка по длине звонка";
            payingParagrapOBL.set_Style("Обычный");
            payingRangeOBL.InsertParagraphAfter();
            var queryOBL =
            from p in dataEntities.Payings
            orderby p.CallLength
            select new { p.Id, p.LastName, p.Phone, p.Date, p.Rate, p.Discount, p.TimeIn, p.TimeOut, p.CallLength };
            var OBL = queryOBL.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapOBL = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeOBL = tableParagrapOBL.Range;
            Microsoft.Office.Interop.Word.Table payingTableOBL = document.Tables.Add(tableRangeOBL, OBL.Count() + 1, 8);
            payingTableOBL.Borders.InsideLineStyle = payingTableOBL.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableOBL.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableOBL.Borders.InsideLineStyle = payingTableOBL.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableOBL.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeOBL;
            cellRangeOBL = payingTableOBL.Cell(1, 1).Range;
            cellRangeOBL.Text = "Фамилия";
            cellRangeOBL = payingTableOBL.Cell(1, 2).Range;
            cellRangeOBL.Text = "Дата оплаты";
            cellRangeOBL = payingTableOBL.Cell(1, 3).Range;
            cellRangeOBL.Text = "Телефон";
            cellRangeOBL = payingTableOBL.Cell(1, 4).Range;
            cellRangeOBL.Text = "Тариф";
            cellRangeOBL = payingTableOBL.Cell(1, 5).Range;
            cellRangeOBL.Text = "Скидка";
            cellRangeOBL = payingTableOBL.Cell(1, 6).Range;
            cellRangeOBL.Text = "Время начала разговора";
            cellRangeOBL = payingTableOBL.Cell(1, 7).Range;
            cellRangeOBL.Text = "Время конца разговора";
            cellRangeOBL = payingTableOBL.Cell(1, 8).Range;
            cellRangeOBL.Text = "Длина разговора";
            for (int i = 0; i < OBL.Count(); ++i)
            {
                var currentphone = OBL.ElementAt(i);
                cellRangeOBL = payingTableOBL.Cell(i + 2, 1).Range;
                cellRangeOBL.Text = currentphone.LastName;
                cellRangeOBL = payingTableOBL.Cell(i + 2, 2).Range;
                cellRangeOBL.Text = currentphone.Date.ToString();
                cellRangeOBL = payingTableOBL.Cell(i + 2, 3).Range;
                cellRangeOBL.Text = currentphone.Phone;
                cellRangeOBL = payingTableOBL.Cell(i + 2, 4).Range;
                cellRangeOBL.Text = currentphone.Rate.ToString();
                cellRangeOBL = payingTableOBL.Cell(i + 2, 5).Range;
                cellRangeOBL.Text = currentphone.Discount.ToString() + "%";
                cellRangeOBL = payingTableOBL.Cell(i + 2, 6).Range;
                cellRangeOBL.Text = currentphone.TimeIn.ToString();
                cellRangeOBL = payingTableOBL.Cell(i + 2, 7).Range;
                cellRangeOBL.Text = currentphone.TimeOut.ToString();
                cellRangeOBL = payingTableOBL.Cell(i + 2, 8).Range;
                cellRangeOBL.Text = currentphone.CallLength.ToString();
            }

            var queryMaxR = dataEntities.Payings.OrderByDescending(i => (i.Rate - (i.Rate / 100 * i.Discount))).FirstOrDefault();
            Microsoft.Office.Interop.Word.Paragraph payingParagrapMaxR = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeMaxR = payingParagrapMaxR.Range;
            payingRangeMaxR.Text = "\nМаксимальный тариф с учетом скидки: " + queryMaxR.Rate.ToString();
            payingParagrapMaxR.set_Style("Обычный");
            payingRangeMaxR.InsertParagraphAfter();

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGLN = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGLN = payingParagrapGLN.Range;
            payingRangeGLN.Text = "Группировка по фамилии";
            payingParagrapGLN.set_Style("Обычный");
            payingRangeGLN.InsertParagraphAfter();
            var queryGLN = dataEntities.Payings.GroupBy(u => u.LastName);
            var GLN = queryGLN.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGLN = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGLN = tableParagrapGLN.Range;
            Microsoft.Office.Interop.Word.Table payingTableGLN = document.Tables.Add(tableRangeGLN, GLN.Count() + 1, 1);
            payingTableGLN.Borders.InsideLineStyle = payingTableGLN.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGLN.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGLN.Borders.InsideLineStyle = payingTableGLN.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGLN.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGLN;
            cellRangeGLN = payingTableGLN.Cell(1, 1).Range;
            cellRangeGLN.Text = "Фамилия";
            for (int i = 0; i < GLN.Count(); ++i)
            {
                cellRangeGLN = payingTableGLN.Cell(i + 2, 1).Range;
                cellRangeGLN.Text = GLN.ElementAt(i).Key;
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGT = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGT = payingParagrapGT.Range;
            payingRangeGT.Text = "Группировка по телефону";
            payingParagrapGT.set_Style("Обычный");
            payingRangeGT.InsertParagraphAfter();
            var queryGT = dataEntities.Payings.GroupBy(u => u.Phone);
            var GT = queryGT.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGT = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGT = tableParagrapGT.Range;
            Microsoft.Office.Interop.Word.Table payingTableGT = document.Tables.Add(tableRangeGT, GT.Count() + 1, 1);
            payingTableGT.Borders.InsideLineStyle = payingTableGT.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGT.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGT.Borders.InsideLineStyle = payingTableGT.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGT.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGT;
            cellRangeGT = payingTableGT.Cell(1, 1).Range;
            cellRangeGT.Text = "Телефон";
            for (int i = 0; i < GT.Count(); ++i)
            {
                cellRangeGT = payingTableGT.Cell(i + 2, 1).Range;
                cellRangeGT.Text = GT.ElementAt(i).Key;
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGD = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGD = payingParagrapGD.Range;
            payingRangeGD.Text = "Группировка по дате оплаты";
            payingParagrapGD.set_Style("Обычный");
            payingRangeGD.InsertParagraphAfter();
            var queryGD = dataEntities.Payings.GroupBy(u => u.Date);
            var GD = queryGD.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGD = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGD = tableParagrapGD.Range;
            Microsoft.Office.Interop.Word.Table payingTableGD = document.Tables.Add(tableRangeGD, GD.Count() + 1, 1);
            payingTableGD.Borders.InsideLineStyle = payingTableGD.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGD.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGD.Borders.InsideLineStyle = payingTableGD.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGD.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGD;
            cellRangeGD = payingTableGD.Cell(1, 1).Range;
            cellRangeGD.Text = "Дата оплаты";
            for (int i = 0; i < GD.Count(); ++i)
            {
                cellRangeGD = payingTableGD.Cell(i + 2, 1).Range;
                cellRangeGD.Text = GD.ElementAt(i).Key.ToString();
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGR = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGR = payingParagrapGR.Range;
            payingRangeGR.Text = "Группировка по тарифу";
            payingParagrapGR.set_Style("Обычный");
            payingRangeGR.InsertParagraphAfter();
            var queryGR = dataEntities.Payings.GroupBy(u => u.Rate);
            var GR = queryGR.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGR = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGR = tableParagrapGR.Range;
            Microsoft.Office.Interop.Word.Table payingTableGR = document.Tables.Add(tableRangeGR, GR.Count() + 1, 1);
            payingTableGR.Borders.InsideLineStyle = payingTableGR.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGR.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGR.Borders.InsideLineStyle = payingTableGR.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGR.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGR;
            cellRangeGR = payingTableGR.Cell(1, 1).Range;
            cellRangeGR.Text = "Тариф";
            for (int i = 0; i < GR.Count(); ++i)
            {
                cellRangeGR = payingTableGR.Cell(i + 2, 1).Range;
                cellRangeGR.Text = GR.ElementAt(i).Key.ToString();
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGDis = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGDis = payingParagrapGDis.Range;
            payingRangeGDis.Text = "Группировка скидке";
            payingParagrapGDis.set_Style("Обычный");
            payingRangeGDis.InsertParagraphAfter();
            var queryGDis = dataEntities.Payings.GroupBy(u => u.Discount);
            var GDis = queryGDis.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGDis = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGDis = tableParagrapGDis.Range;
            Microsoft.Office.Interop.Word.Table payingTableGDis = document.Tables.Add(tableRangeGDis, GDis.Count() + 1, 1);
            payingTableGDis.Borders.InsideLineStyle = payingTableGDis.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGDis.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGDis.Borders.InsideLineStyle = payingTableGDis.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGDis.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGDis;
            cellRangeGDis = payingTableGDis.Cell(1, 1).Range;
            cellRangeGDis.Text = "Скидка";
            for (int i = 0; i < GDis.Count(); ++i)
            {
                cellRangeGDis = payingTableGDis.Cell(i + 2, 1).Range;
                cellRangeGDis.Text = GDis.ElementAt(i).Key.ToString()+"%";
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGTI = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGTI = payingParagrapGTI.Range;
            payingRangeGTI.Text = "Группировка по времени начала разговора";
            payingParagrapGTI.set_Style("Обычный");
            payingRangeGTI.InsertParagraphAfter();
            var queryGTI = dataEntities.Payings.GroupBy(u => u.TimeIn);
            var GTI = queryGTI.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGTI = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGTI = tableParagrapGTI.Range;
            Microsoft.Office.Interop.Word.Table payingTableGTI = document.Tables.Add(tableRangeGTI, GTI.Count() + 1, 1);
            payingTableGTI.Borders.InsideLineStyle = payingTableGTI.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGTI.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGTI.Borders.InsideLineStyle = payingTableGTI.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGTI.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGTI;
            cellRangeGTI = payingTableGTI.Cell(1, 1).Range;
            cellRangeGTI.Text = "Время начала разговора";
            for (int i = 0; i < GTI.Count(); ++i)
            {
                cellRangeGTI = payingTableGTI.Cell(i + 2, 1).Range;
                cellRangeGTI.Text = GTI.ElementAt(i).Key.ToString();
            }

            Microsoft.Office.Interop.Word.Paragraph payingParagrapGTO = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range payingRangeGTO = payingParagrapGTO.Range;
            payingRangeGTO.Text = "Группировка по времени конца разговора";
            payingParagrapGTO.set_Style("Обычный");
            payingRangeGTO.InsertParagraphAfter();
            var queryGTO = dataEntities.Payings.GroupBy(u => u.TimeOut);
            var GTO = queryGTO.ToList();
            Microsoft.Office.Interop.Word.Paragraph tableParagrapGTO = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tableRangeGTO = tableParagrapGTO.Range;
            Microsoft.Office.Interop.Word.Table payingTableGTO = document.Tables.Add(tableRangeGTO, GTO.Count() + 1, 1);
            payingTableGTO.Borders.InsideLineStyle = payingTableGTO.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGTO.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            payingTableGTO.Borders.InsideLineStyle = payingTableGTO.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            payingTableGTO.Range.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            Microsoft.Office.Interop.Word.Range cellRangeGTO;
            cellRangeGTO = payingTableGTO.Cell(1, 1).Range;
            cellRangeGTO.Text = "Время конца разговора";
            for (int i = 0; i < GTO.Count(); ++i)
            {
                cellRangeGTO = payingTableGTO.Cell(i + 2, 1).Range;
                cellRangeGTO.Text = GTO.ElementAt(i).Key.ToString();
            }

            application.Visible = true;
            document.SaveAs2(@"C:\Users\kapoo\OneDrive\Рабочий стол");
        }

        
    }
}
