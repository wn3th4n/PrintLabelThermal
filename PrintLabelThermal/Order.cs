using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintLabelThermal
{
    internal class Order
    {
        public string OrderID { get; set; }
        public string Date { get; set; }
        public string OrderName { get; set; }
        public int Quantity { get; set; }
        public string Price { get; set; }

        public override string ToString()
        {
            return $"OrderID: {OrderID}, Date: {Date}, Order: {OrderName}, Quantity: {Quantity}, Price: {Price}";
        }

    }
}
