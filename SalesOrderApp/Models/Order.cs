using System;
using System.Collections.Generic;

namespace SalesOrderApp.Models
{
    public class Order
    {
        public int OrderId { get; set; }
        public DateTime Date { get; set; }
        public int CustomerId { get; set; }
        public Customer Customer { get; set; }
        public decimal NetAmount { get; set; }
        public List<OrderDetail> OrderDetails { get; set; }
    }
}
