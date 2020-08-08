using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNB_Excel_Add_In
{
    class CurrencyData
    {
        public int Unit { get; set; }
        public string Currency { get; set; }
        public double Rate { get; set; }
        public CurrencyData(int unit, string currency, double rate)
        {
            this.Unit = unit;
            this.Currency = currency;
            this.Rate = rate;
        }
    }
}
