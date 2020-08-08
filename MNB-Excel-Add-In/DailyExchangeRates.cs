using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNB_Excel_Add_In
{
    class DailyExchangeRates
    {
        public string DateOfExchangeRate { get; set; }
        public List<CurrencyData> CurrencyData { get; set; }

        public DailyExchangeRates(string day, List<CurrencyData> currencyData)
        {
            this.DateOfExchangeRate = day;
            this.CurrencyData = currencyData;
        }
        public override string ToString()
        {
            return this.DateOfExchangeRate + "\n" + string.Join("\n", CurrencyData.Select(x => x.Currency + "\t" + x.Rate + "\t" + x.Unit).ToArray());
        }
    }
}
