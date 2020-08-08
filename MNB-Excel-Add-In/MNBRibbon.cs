using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Tools.Ribbon;
using MNB_Excel_Add_In.hu.mnb.www;

namespace MNB_Excel_Add_In
{
    public partial class MNBRibbon
    {
        List<string> currencies;
        string startdate = "2020-03-01";
        string enddate = "2020-03-31";

        string Currencies { 
            get{
                return currencies == null ? "" : string.Join(",", currencies);
            } 
        }

        private void MNBRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            currencies = new List<string>();
        }

        private void mnbDataBTN_Click(object sender, RibbonControlEventArgs e)
        {
            currencies = GetCurrencyTypesFromWebservice();
            var currencyUnitDictionary = GetUnitForTypesFromWebservice(Currencies);
            GetCurrencyRatesInInterval(startdate, enddate);
        }

        List<string> GetCurrencyTypesFromWebservice()
        {
            string currResponseResult = "";
            List<string> currencies = new List<string>();

            using (MNBArfolyamServiceSoapImpl test = new MNBArfolyamServiceSoapImpl())
            {
                var result = test.GetCurrencies(new GetCurrenciesRequestBody());
                currResponseResult = result.GetCurrenciesResult;
            }

            XmlDocument currResponse = new XmlDocument();
            currResponse.LoadXml(currResponseResult);
            XmlNodeList currs = currResponse.SelectNodes("/MNBCurrencies/Currencies/Curr");

            
            foreach (XmlNode curr in currs)
                currencies.Add(curr.InnerText);

            return currencies;
        }

        Dictionary<string, int> GetUnitForTypesFromWebservice(string currencies)
        {
            Dictionary<string, int> currencyUnits = new Dictionary<string, int>();
            string currResponseResult = "";

            using (MNBArfolyamServiceSoapImpl test = new MNBArfolyamServiceSoapImpl())
            {
                var result = test.GetCurrencyUnits(new GetCurrencyUnitsRequestBody() { currencyNames = Currencies });
                currResponseResult = result.GetCurrencyUnitsResult;
            }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(currResponseResult);
            XmlNodeList units = xmlDocument.SelectNodes("/MNBCurrencyUnits/Units/Unit");

            foreach (XmlNode unit in units)
            {
                var dayOfExchangeRate = unit.Attributes["curr"].Value;
                var unitOfCurr = int.Parse(unit.InnerText);
                currencyUnits.Add(dayOfExchangeRate, unitOfCurr);
            }

            return currencyUnits;
        }

        //have to check if the date is valid
        List<DailyExchangeRates> GetCurrencyRatesInInterval(string startdate,string enddate)
        {
            GetExchangeRatesResponseBody result;
            using (MNBArfolyamServiceSoapImpl test = new MNBArfolyamServiceSoapImpl())
            {
                //<MNBExchangeRates><Day date="2020-04-01"><Rate unit="1" curr="EUR">364,57</Rate></Day></MNBExchangeRates>
                var myExchangeratesRequestBody = new GetExchangeRatesRequestBody() { startDate = startdate, endDate = enddate, currencyNames = Currencies };
                result = test.GetExchangeRates(myExchangeratesRequestBody);

            }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(result.GetExchangeRatesResult);
            XmlNodeList xnList = xmlDocument.SelectNodes("/MNBExchangeRates/Day");
            List<DailyExchangeRates> dailyRates = new List<DailyExchangeRates>();
            foreach (XmlNode day in xnList)
            {
                List<CurrencyData> currencyData = new List<CurrencyData>();
                var dayOfExchangeRate = day.Attributes["date"].Value;
                var dailyCurrencyExchangeRates = day.SelectNodes("Rate");
                foreach (XmlNode exchangeRates in dailyCurrencyExchangeRates)
                {
                    int unit = int.Parse(exchangeRates.Attributes["unit"].Value);
                    string curr = exchangeRates.Attributes["curr"].Value;
                    double value = double.Parse(exchangeRates.InnerText);
                    currencyData.Add(new CurrencyData(unit, curr, value));
                    //Console.WriteLine(exchangeRates.Attributes["curr"].Value + "\t" + exchangeRates.InnerText);
                }
                dailyRates.Add(new DailyExchangeRates(dayOfExchangeRate, currencyData));
            }
            foreach (var day in dailyRates)
            {
                Console.WriteLine(day.ToString());
            }
            return dailyRates;
        }

    }
}
