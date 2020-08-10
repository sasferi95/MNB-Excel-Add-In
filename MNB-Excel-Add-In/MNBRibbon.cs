using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using MNB_Excel_Add_In.hu.mnb.www;
using DataTable = System.Data.DataTable;

namespace MNB_Excel_Add_In
{
    public partial class MNBRibbon
    {
        const string OleDbConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Resources\ExcelButton.accdb";
        List<string> currencies;
        string startdate = "";
        string enddate = "";

        string Currencies { 
            get{
                return currencies == null ? "" : string.Join(",", currencies);
            } 
        }

        private void MNBRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            currencies = new List<string>();
            startdate = "2015-01-01";
            enddate = "2020-04-01";
        }

        /// <summary>
        /// Registers the click event on the 'mnb adatletöltés button
        /// (it is a task to not freeze the ui of the excel)
        /// </summary>
        /// <param name="sender">sender button</param>
        /// <param name="e">event args</param>
        private async void mnbDataBTN_Click(object sender, RibbonControlEventArgs e)
        {
            //Task t= new Task(() => FillData());
            //try
            //{
            //    t.Start();
            //    await t;
            //}
            //catch (Exception exc)
            //{
            //    MessageBox.Show("Error",exc.Message,MessageBoxButtons.OK);
            //}


            //I decided to go with sequential solution which blocks the gui till its finished
            DoWork();
        }

        /// <summary>
        /// 'filler' method to be able to create a task
        /// </summary>
        void DoWork()
        {
            currencies = GetCurrencyTypesFromWebservice();

            //dictionary for currency and its unit
            var currencyUnitDictionary = GetUnitForTypesFromWebservice(Currencies);

            //get exchange rates in a given timespan but weekends excluded?!
            var dailyExchangeRates = GetCurrencyRatesInInterval(startdate, enddate);
            var orderedDailyExchangeRates = dailyExchangeRates.OrderBy(x => x.DateOfExchangeRate).ToList();

            var offsetCurrencyColumn = OffsetDictionaryMaker(currencies, 2);
            var offsetDatesRow = OffsetDictionaryMaker(orderedDailyExchangeRates.Select(x => x.DateOfExchangeRate).ToList(), 3);
            
            var activeSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            InsertExcelCurrencyHeader(activeSheet, currencyUnitDictionary, offsetCurrencyColumn);       
            InsertExcelCurrencyRatesWithDates(activeSheet, dailyExchangeRates, offsetCurrencyColumn, offsetDatesRow);

            string username = GetUser();
            var timestamp = DateTime.Now;

            InsertNewLog(username, timestamp);
        }

        /// <summary>
        /// gets the curreny types from webserice
        /// </summary>
        /// <returns>returns the list of currencies</returns>
        List<string> GetCurrencyTypesFromWebservice()
        {
            string currResponseResult = "";
            List<string> currencies = new List<string>();

            try
            {
                using (MNBArfolyamServiceSoapImpl test = new MNBArfolyamServiceSoapImpl())
                {
                    var result = test.GetCurrencies(new GetCurrenciesRequestBody());
                    currResponseResult = result.GetCurrenciesResult;
                }
            }
            catch(Exception e)
            {
                throw new Exception("Error happened while getting currency types from webservice.\n" + e.Message, e);
            }

            XmlDocument currResponse = new XmlDocument();
            currResponse.LoadXml(currResponseResult);
            XmlNodeList currs = currResponse.SelectNodes("/MNBCurrencies/Currencies/Curr");

            foreach (XmlNode curr in currs)
                currencies.Add(curr.InnerText);

            return currencies;
        }

        /// <summary>
        /// Gets the currencies' unit and creates a dictionary from it
        /// </summary>
        /// <param name="currencies">all currency seperated by a ,</param>
        /// <returns>dictionary for each currency and the currency's unit</returns>
        Dictionary<string, int> GetUnitForTypesFromWebservice(string currencies)
        {
            Dictionary<string, int> currencyUnits = new Dictionary<string, int>();
            string currResponseResult = "";

            try
            {
                using (MNBArfolyamServiceSoapImpl mnbService = new MNBArfolyamServiceSoapImpl())
                {
                    var result = mnbService.GetCurrencyUnits(new GetCurrencyUnitsRequestBody() { currencyNames = Currencies });
                    currResponseResult = result.GetCurrencyUnitsResult;
                }
            }
            catch(Exception e)
            {
                throw new Exception("Error happened while fetching currencies' units.\n" + e.Message, e);
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

        /// <summary>
        /// gets the currency rates between two dates
        /// </summary>
        /// <param name="startdate">dirst day of the interval</param>
        /// <param name="enddate">last day of the interval</param>
        /// <returns></returns>
        List<DailyExchangeRates> GetCurrencyRatesInInterval(string startdate,string enddate)
        {
            GetExchangeRatesResponseBody result;
            try
            {
                using (MNBArfolyamServiceSoapImpl mnbService = new MNBArfolyamServiceSoapImpl())
                {
                    //<MNBExchangeRates><Day date="2020-04-01"><Rate unit="1" curr="EUR">364,57</Rate></Day></MNBExchangeRates>
                    var myExchangeratesRequestBody = new GetExchangeRatesRequestBody() { startDate = startdate, endDate = enddate, currencyNames = Currencies };
                    result = mnbService.GetExchangeRates(myExchangeratesRequestBody);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error happened while fetching exchange rates from webservice.\n" + e.Message, e);
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
                }
                dailyRates.Add(new DailyExchangeRates(dayOfExchangeRate, currencyData));
            }
            foreach (var day in dailyRates)
            {
                Console.WriteLine(day.ToString());
            }
            return dailyRates;
        }

        /// <summary>
        /// creates the 'header' part of the ecel file with the currencies and units 
        /// </summary>
        /// <param name="activeSheet">the active worksheet</param>
        /// <param name="currencyUnits">currency -> unit dictionary</param>
        /// <param name="offsetCurrencyColumn">help to determine which column which currency's</param>
        void InsertExcelCurrencyHeader(Worksheet activeSheet, Dictionary<string,int> currencyUnits,Dictionary<string, int> offsetCurrencyColumn)
        {
            activeSheet.Cells[1, 1].Value2 = "Dátum/ISO";
            activeSheet.Cells[2, 1].Value2 = "Egység";

            foreach(KeyValuePair<string, int> keyValuePair in currencyUnits)
            {
                string curCurrency = keyValuePair.Key;
                int column = offsetCurrencyColumn[curCurrency];
                activeSheet.Cells[1, column].Value2 = keyValuePair.Key;
                activeSheet.Cells[2, column].Value2 = keyValuePair.Value;
                activeSheet.Cells[2, column].NumberFormat = "0";
            }
        }

        /// <summary>
        /// inserts the days and the exchange rates into the excel
        /// </summary>
        /// <param name="activeSheet">the active worksheet</param>
        /// <param name="dailyExchangeRates">stores each days exchange rates</param>
        /// <param name="currencyForColumn">offset dictionary to help position the rates' columns</param>
        /// <param name="dateForRow">offset dictionary to help position the rates' rows</param>
        void InsertExcelCurrencyRatesWithDates(Worksheet activeSheet, List<DailyExchangeRates> dailyExchangeRates, Dictionary<string, int> currencyForColumn, Dictionary<string, int> dateForRow)
        {
            try
            {
                foreach (var daylyRates in dailyExchangeRates)
                {
                    int row = dateForRow[daylyRates.DateOfExchangeRate];
                    activeSheet.Cells[row, 1].Value2 = daylyRates.DateOfExchangeRate;
                    //Globals.ThisAddIn.Application.ActiveSheet.Cells[row, 1].NumberFormat = "YYYY.mm.hh.";
                    foreach (CurrencyData cd in daylyRates.CurrencyDatas)
                    {
                        int col = currencyForColumn[cd.Currency];
                        activeSheet.Cells[row, col].Value2 = cd.Rate;
                        //Globals.ThisAddIn.Application.ActiveSheet.Cells[row, col].NumberFormat = "0,00";
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error happened during inserting data to the excel.\n" + e.Message, e);
            }
            if (dailyExchangeRates.Count != 0)
            {
                int firstCol = currencyForColumn.Values.Min();
                int lastCol = currencyForColumn.Values.Max();
                int firstRow = dateForRow.Values.Min();
                int lastRow = dateForRow.Values.Max();
                activeSheet.Range[activeSheet.Cells[firstRow, firstCol], activeSheet.Cells[lastRow, lastCol]].NumberFormat = "0,00";
            }
        }

        /// <summary>
        /// helper function to create offsets for positioning
        /// </summary>
        /// <param name="list">input list</param>
        /// <param name="start">starting row/col</param>
        /// <returns>dictionary with the list elements as keys and its offset</returns>
        Dictionary<string, int> OffsetDictionaryMaker(List<string> list, int start = 0)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            list.ForEach(x => dict.Add(x, start++));
            return dict;
        }

        /// <summary>
        /// In case of log button have been pressed opens a windows to display the button presses and allows the modification if comments for each log
        /// </summary>
        /// <param name="sender">button</param>
        /// <param name="e">event args</param>
        private void logBtn_Click(object sender, RibbonControlEventArgs e)
        {
            new LogWindow().ShowDialog();
        }

        /// <summary>
        /// get user together with the domain
        /// </summary>
        /// <returns>domainname//username as a string</returns>
        private string GetUser()
        {
            return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        }

        /// <summary>
        /// when fetched data from webservice a new log is created if the user successfully received id and the document is ready
        /// </summary>
        /// <param name="username">domainname/username</param>
        /// <param name="logTime">time when button pressed</param>
        //raise error
        void InsertNewLog(string username, DateTime logTime)
        {
            //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Resources\ExcelButton.accdb
            using (OleDbConnection myConn = new OleDbConnection(OleDbConnectionString))
            {                
                string query = "INSERT INTO MNBButtonLogs ([DomainUsername], [Timestamp]) " +
                               "VALUES (@username, @timestamp)";

                OleDbCommand mySelectCommand = new OleDbCommand(query, myConn);

                mySelectCommand.Parameters.AddWithValue("@username", username);
                mySelectCommand.Parameters.AddWithValue("@timestamp", logTime.ToString("yyyy-MM-dd HH:mm:ss"));

                try
                {
                    myConn.Open();

                    mySelectCommand.ExecuteNonQuery();
                }
                catch(Exception e)
                {
                    throw new Exception("Could not log the button press.\n" + e.Message, e);
                }
            }
        }
    }
}
