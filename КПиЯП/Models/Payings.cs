//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace КПиЯП.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Payings
    {
        public int Id { get; set; }
        public string LastName { get; set; }
        public string Phone { get; set; }
        public System.DateTime Date { get; set; }
        public decimal Rate { get; set; }
        public int Discount { get; set; }
        public System.TimeSpan TimeIn { get; set; }
        public System.TimeSpan TimeOut { get; set; }
        public System.TimeSpan CallLength { get; set; }
    }
}
