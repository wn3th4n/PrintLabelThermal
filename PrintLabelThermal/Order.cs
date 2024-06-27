using Newtonsoft.Json;
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
        [JsonConverter(typeof(StringArrayConvert))]
        public string[] Orders { get; set; }
        [JsonConverter(typeof(StringArrayConvert))]
        public string[] Notes { get; set; }
        public string TotalPrice { get; set; }

   

    }
}
