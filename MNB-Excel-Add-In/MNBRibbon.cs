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

    }
}
