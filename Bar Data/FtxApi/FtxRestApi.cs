using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using FtxApi.Enums;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;


namespace FtxApi
{
    public class FtxRestApi
    {
        private const string Url = "https://ftx.com/";

        private readonly Client _client;

        private readonly HttpClient _httpClient;

        private readonly HMACSHA256 _hashMaker;

        private long _nonce;

        // Self-initiated class 

        //Helper function for INT ( UNIX TIME-STAMP TO Date Time)

        public static DateTime UnixTimeStampToDateTime(int unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp).ToLocalTime();
            return dateTime;
        }

        // Orderbook (Bids and Asks)

        public class JSONModel
        {
            public class Orderbook
            {
                public bool success { get; set; }
                public Result result { get; set; }
            }

            public class Result
            {
                public float[][] bids { get; set; }
                public float[][] asks { get; set; }
            }

        }

        // Account detail ( balance + position , etc)
        public class account
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result result { get; set; }
            }

            public class Result
            {
                public int accountIdentifier { get; set; }
                public string username { get; set; }
                public float collateral { get; set; }
                public float freeCollateral { get; set; }
                public float totalAccountValue { get; set; }
                public float totalPositionSize { get; set; }
                public float initialMarginRequirement { get; set; }
                public float maintenanceMarginRequirement { get; set; }
                public object marginFraction { get; set; }
                public object openMarginFraction { get; set; }
                public bool liquidating { get; set; }
                public bool backstopProvider { get; set; }
                public Position[] positions { get; set; }
                public float takerFee { get; set; }
                public float makerFee { get; set; }
                public float leverage { get; set; }
                public float futuresLeverage { get; set; }
                public object positionLimit { get; set; }
                public object positionLimitUsed { get; set; }
                public bool useFttCollateral { get; set; }
                public bool chargeInterestOnNegativeUsd { get; set; }
                public bool spotMarginEnabled { get; set; }
                public bool spotMarginWithdrawalsEnabled { get; set; }
                public bool spotLendingEnabled { get; set; }
                public object accountType { get; set; }
            }

            public class Position
            {
                public string future { get; set; }
                public float size { get; set; }
                public string side { get; set; }
                public float netSize { get; set; }
                public float longOrderSize { get; set; }
                public float shortOrderSize { get; set; }
                public float cost { get; set; }
                public object entryPrice { get; set; }
                public float unrealizedPnl { get; set; }
                public float realizedPnl { get; set; }
                public float initialMarginRequirement { get; set; }
                public float maintenanceMarginRequirement { get; set; }
                public float openSize { get; set; }
                public float collateralUsed { get; set; }
                public object estimatedLiquidationPrice { get; set; }
            }
        }

        //


        public FtxRestApi(Client client)
        {
            _client = client;
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(Url),
                Timeout = TimeSpan.FromSeconds(30)
            };

            _hashMaker = new HMACSHA256(Encoding.UTF8.GetBytes(_client.ApiSecret));
        }

        #region Coins

        public async Task<dynamic> GetCoinsAsync()
        {
            var resultString = $"api/coins";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        #endregion

        #region Futures

        public async Task<dynamic> GetAllFuturesAsync()
        {
            var resultString = $"api/futures";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetFutureAsync(string future)
        {
            var resultString = $"api/futures/{future}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetFutureStatsAsync(string future)
        {
            var resultString = $"api/futures/{future}/stats";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetFundingRatesAsync(DateTime start, DateTime end)
        {
            var resultString = $"api/funding_rates?start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetHistoricalPricesAsync(string futureName, int resolution, int limit, DateTime start, DateTime end)
        {
            var resultString = $"api/futures/{futureName}/mark_candles?resolution={resolution}&limit={limit}&start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetHistoricalSpotPricesAsync(string SpotName, int resolution, DateTime start, DateTime end)
        {
            var resultString = $"api/markets/{SpotName}/candles?resolution={resolution}&start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        #endregion

        #region Markets

        public async Task<dynamic> GetMarketsAsync()
        {
            var resultString = $"api/markets";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetSingleMarketsAsync(string marketName)
        {
            var resultString = $"api/markets/{marketName}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetMarketOrderBookAsync(string marketName, int depth = 20)
        {
            var resultString = $"api/markets/{marketName}/orderbook?depth={depth}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }


        public async Task<dynamic> GetMarketTradesAsync(string marketName, int limit, DateTime start, DateTime end)
        {
            var resultString = $"api/markets/{marketName}/trades?limit={limit}&start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }



        #endregion

        #region Account

        public async Task<dynamic> GetAccountInfoAsync()
        {
            var resultString = $"api/account";
            var sign = GenerateSignature(HttpMethod.Get, "/api/account", "");
            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetPositionsAsync()
        {
            var resultString = $"api/positions";
            var sign = GenerateSignature(HttpMethod.Get, "/api/positions", "");
            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> ChangeAccountLeverageAsync(int leverage)
        {
            var resultString = $"api/account/leverage";

            var body = $"{{\"leverage\": {leverage}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/account/leverage", body);

            var result = await CallAsyncSign(HttpMethod.Post, resultString, sign, body);

            return ParseResponce(result);
        }
        // Self-defined


        public class SnapPoint
        {
            public bool success { get; set; }
            public int result { get; set; }
        }



        public async Task<dynamic> RequestHistoricalBalance(string[] accounts, int endTime)
        {
            var path = $"api/historical_balances/requests";

            string accountsStr = "[\"" + string.Join("\",\"", accounts) + "\"]";

            var body = $"{{" +
               $"\"accounts\": {accountsStr}," +
               $"\"endTime\": {endTime}" +
               "}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/historical_balances/requests", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);
            return ParseResponce(result);
        }


        public int Get_SnapPoint(string[] accounts, int endTime)
        {
            SnapPoint records = JsonConvert.DeserializeObject<SnapPoint>(RequestHistoricalBalance(accounts, endTime).Result);

            return records.result;

        }
        // Self-written 


        public class SubAccount
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public string nickname { get; set; }
                public bool deletable { get; set; }
                public bool editable { get; set; }
                public bool competition { get; set; }
            }


        }

        public async Task<dynamic> GetSubaccounts()
        {
            var resultString = $"api/subaccounts";
            var sign = GenerateSignature(HttpMethod.Get, "/api/subaccounts", "");
            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }


        public string[] GetAccountsList()
        {

            SubAccount.Rootobject records = JsonConvert.DeserializeObject<SubAccount.Rootobject>(GetSubaccounts().Result);
            string[] result = new string[records.result.Length + 1];
            result[0] = "main";
            for (int i = 1; i < records.result.Length + 1; i++)
            {
                result[i] = records.result[i].nickname;
            }

            return result;
        }

        public async Task<dynamic> GetHistoricalBalances()
        {
            var resultString = $"api/historical_balances/requests";
            var sign = GenerateSignature(HttpMethod.Get, "/api/historical_balances/requests", "");
            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetHistoricalBalancesWith_ID(int request_id)
        {
            var resultString = $"api/historical_balances/requests/{request_id.ToString()}";
            var sign = GenerateSignature(HttpMethod.Get, "/api/historical_balances/requests/", request_id.ToString());
            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> CreateSubaccounts(string name)
        {
            var path = $"api/subaccounts";

            var body = $"{{\"nickname\": \"{name}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/subaccounts", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);
            return ParseResponce(result);
        }

        //
        #endregion

        #region Wallet

        public async Task<dynamic> GetCoinAsync()
        {
            var resultString = $"api/wallet/coins";

            var sign = GenerateSignature(HttpMethod.Get, "/api/wallet/coins", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetBalancesAsync()
        {
            var resultString = $"api/wallet/balances";

            var sign = GenerateSignature(HttpMethod.Get, "/api/wallet/balances", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetDepositAddressAsync(string coin)
        {
            var resultString = $"api/wallet/deposit_address/{coin}";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/wallet/deposit_address/{coin}", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetDepositHistoryAsync()
        {
            var resultString = $"api/wallet/deposits";

            var sign = GenerateSignature(HttpMethod.Get, "/api/wallet/deposits", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetWithdrawalHistoryAsync()
        {
            var resultString = $"api/wallet/withdrawals";

            var sign = GenerateSignature(HttpMethod.Get, "/api/wallet/withdrawals", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> RequestWithdrawalAsync(string coin, decimal size, string addr, string tag, string pass, string code)
        {
            var resultString = $"api/wallet/withdrawals";

            var body = $"{{" +
                $"\"coin\": \"{coin}\"," +
                $"\"size\": {size}," +
                $"\"address\": \"{addr}\"," +
                $"\"tag\": {tag}," +
                $"\"password\": \"{pass}\"," +
                $"\"code\": {code}" +
                "}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/wallet/withdrawals", body);

            var result = await CallAsyncSign(HttpMethod.Post, resultString, sign, body);

            return ParseResponce(result);
        }

        // Self-written Functions for find the closeest bid price in the orderbook
        // Book Order
        // Turn it from string ==> Json ==> Class and object

        // Balance
        // First input the type of Coin you want to search, e.g. USD, PERP, BTC, ETH,
        // it will then show you the total amount of that coin
        // You can also use GetCoinBalance_value to check the value in "USD"

        public class Balance
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public string coin { get; set; }
                public float total { get; set; }
                public float free { get; set; }
                public float availableForWithdrawal { get; set; }
                public float availableWithoutBorrow { get; set; }
                public float usdValue { get; set; }
                public float spotBorrow { get; set; }
            }

        }

        public float GetCoinBalance(string type)
        {
            Balance.Rootobject records = JsonConvert.DeserializeObject<Balance.Rootobject>(GetBalancesAsync().Result);

            for (int i = 0; i <= records.result.Length - 1; i++)
            {
                if (string.Equals(type, records.result[i].coin))
                {
                    return records.result[i].total;
                };
            }

            return 0;
        }



        public float GetCoinBalance_Value(string type)
        {
            Balance.Rootobject records = JsonConvert.DeserializeObject<Balance.Rootobject>(GetBalancesAsync().Result);

            for (int i = 0; i <= records.result.Length - 1; i++)
            {
                if (string.Equals(type, records.result[i].coin))
                {
                    return records.result[i].usdValue;
                };
            }

            return 0;
        }



        // End of Balance


        // Bar Data chart

        public class BarData
        {
            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public DateTime startTime { get; set; }
                public float time { get; set; }
                public float open { get; set; }
                public float high { get; set; }
                public float low { get; set; }
                public float close { get; set; }
                public float volume { get; set; }
            }
        }


        // End of Bar data


        /* Task Give by Karen
        1. Account balance
        2. Trade History
        3. Withdrawal and Deposit History
        */

        // 1. Account balance

        //Break Down of Balance
        class Account_Balances_Analysis_Native
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public string coin { get; set; }
                public float total { get; set; }
                public float free { get; set; }
                public float availableForWithdrawal { get; set; }
                public float availableWithoutBorrow { get; set; }
                public float usdValue { get; set; }
                public float spotBorrow { get; set; }
            }

        }

        public class Account_Balances_Analysis
        {

            public string[] coin;
            public float[] total;
            public float[] free;
            public float[] avaliableForWithdrawal;
            public float[] avaliableWithoutBorrow;
            public float[] usdValue;
            public float[] spotBorrow;

            public Account_Balances_Analysis(int length)
            {
                coin = new string[length];
                total = new float[length];
                free = new float[length];
                avaliableForWithdrawal = new float[length];
                avaliableWithoutBorrow = new float[length];
                usdValue = new float[length];
                spotBorrow = new float[length];
            }
        }

        public Account_Balances_Analysis Account_Balance_Analysis()
        {

            Account_Balances_Analysis_Native.Rootobject records = JsonConvert.DeserializeObject<Account_Balances_Analysis_Native.Rootobject>(GetBalancesAsync().Result);

            int length = records.result.Length;

            Account_Balances_Analysis result = new Account_Balances_Analysis(length);

            for (int i = 0; i < records.result.Length; i++)
            {
                result.coin[i] = records.result[i].coin;
                result.total[i] = records.result[i].total;
                result.free[i] = records.result[i].free;
                result.avaliableForWithdrawal[i] = records.result[i].availableForWithdrawal;
                result.avaliableWithoutBorrow[i] = records.result[i].availableWithoutBorrow;
                result.usdValue[i] = records.result[i].usdValue;
                result.spotBorrow[i] = records.result[i].spotBorrow;

            }


            return result;

        }

        // Historical Balance


        public class Historical_Balances_Analysis_Native
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public int id { get; set; }
                public string[] accounts { get; set; }
                public DateTime time { get; set; }
                public DateTime endTime { get; set; }
                public string status { get; set; }
                public bool error { get; set; }
                public Result1[] results { get; set; }
            }

            public class Result1
            {
                public string account { get; set; }
                public string ticker { get; set; }
                public float size { get; set; }
                public float price { get; set; }
            }

        }


        public class Historical_Balances_Analysis
        {
            public DateTime[] endTime;
            public float[] value;

            public Historical_Balances_Analysis(int length)
            {

                endTime = new DateTime[length];
                value = new float[length];

            }

        }

        public Historical_Balances_Analysis GetAnalysis_HistoricalBalances(int[] snapshotList)
        {
            Historical_Balances_Analysis analysis = new Historical_Balances_Analysis(snapshotList.Length);
            for (int i = 0; i < snapshotList.Length; i++)
            {
                Historical_Balances_Analysis_Native.Rootobject records = JsonConvert.DeserializeObject<Historical_Balances_Analysis_Native.Rootobject>(GetBalancesSnapshot(snapshotList[i]).Result);


                analysis.endTime[i] = records.result[i].endTime;
                float sum = 0;
                for (int j = 0; j < records.result[i].results.Length; j++)
                {
                    sum = sum + records.result[i].results[j].size * records.result[i].results[j].price;
                }

                analysis.value[i] = sum;
            }
            return analysis;
        }


        public void Create_Historical_Balances_Excel(Historical_Balances_Analysis analysis, Excel.Application Application)
        {
            if (Application != null)
            {

                Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();

                Worksheet.Cells[1, 1] = "Date";
                Worksheet.Cells[1, 2] = "Value";

                for (int i = 0; i < analysis.endTime.Length; i++)
                {
                    Worksheet.Cells[i + 2, 1] = analysis.endTime[i];
                    Worksheet.Cells[i + 2, 2] = analysis.value[i];

                }
            }
            else
            {
                Console.WriteLine("Historical Data Error");
            };
        }
        //End of Historical Balances


        // Historical Prices

        public class Historical_Price_Analysis
        {

            public DateTime[] time;
            public float[] open;
            public float[] high;
            public float[] low;
            public float[] close;
            public float[] volume;

            public Historical_Price_Analysis(int length)
            {
                time = new DateTime[length];
                open = new float[length];
                high = new float[length];
                low = new float[length];
                close = new float[length];
                volume = new float[length];
            }
        }




        public Historical_Price_Analysis RecordOfHistoricalData(string type, int resolution, int limit, DateTime start, DateTime end)
        {
            BarData.Rootobject records = JsonConvert.DeserializeObject<BarData.Rootobject>(GetHistoricalPricesAsync(type, resolution, limit, start, end).Result);

            int length = records.result.Length;

            Historical_Price_Analysis result = new Historical_Price_Analysis(length);

            for (int i = 0; i < records.result.Length; i++)
            {

                result.time[i] = records.result[i].startTime;
                result.open[i] = records.result[i].open;
                result.high[i] = records.result[i].high;
                result.low[i] = records.result[i].low;
                result.close[i] = records.result[i].close;
                result.volume[i] = records.result[i].volume;
            }


            return result;
        }

        public class Bar_Data_Spot
        {
            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public DateTime startTime { get; set; }
                public float time { get; set; }
                public float open { get; set; }
                public float high { get; set; }
                public float low { get; set; }
                public float close { get; set; }
                public float volume { get; set; }
            }
        }



        public Historical_Price_Analysis RecordOfHistoricalData_Spot(string type, int resolution, DateTime start, DateTime end)
        {
            Bar_Data_Spot.Rootobject records = JsonConvert.DeserializeObject<Bar_Data_Spot.Rootobject>(GetHistoricalSpotPricesAsync(type, resolution, start, end).Result);

            if (records == null) { return null; }

            int length = records.result.Length;

            Historical_Price_Analysis result = new Historical_Price_Analysis(length);

            for (int i = 0; i < records.result.Length; i++)
            {

                result.time[i] = records.result[i].startTime;
                result.open[i] = records.result[i].open;
                result.high[i] = records.result[i].high;
                result.low[i] = records.result[i].low;
                result.close[i] = records.result[i].close;
                result.volume[i] = records.result[i].volume;
            }


            return result;
        }


        // Create an Excel


        public Excel.Application create_Empty_Excel_Application()
        {
            Excel.Application Application = new Excel.Application();
            Application.Workbooks.Add();
            return Application;

        }

        public void Create_New_Workbook(Excel.Application application, Excel.Worksheet worksheet)
        {
            if (application != null)
            {
                Excel.Workbook workbook = application.Workbooks.Add();
                workbook.Worksheets.Add(worksheet);

            }

        }

        public void Save_Excel_File(Excel.Application Application, string SaveLocation)
        {

            if (Application != null)
            {
                Application.ActiveWorkbook.SaveAs(SaveLocation, Excel.XlFileFormat.xlCSV);
                Application.ActiveWorkbook.Close(0);
                Application.Quit();
            }
            else
            {
                Console.WriteLine("Save File Error");
            }

        }

        public void Create_Excel_Breakdown_Balances(Account_Balances_Analysis analysis, Excel.Application Application)
        {
            if (Application != null)
            {

                Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();

                Worksheet.Cells[1, 1] = "Coin";
                Worksheet.Cells[1, 2] = "Total";
                Worksheet.Cells[1, 3] = "Free";
                Worksheet.Cells[1, 4] = "Avaliable For Withdrawal";
                Worksheet.Cells[1, 5] = "Avaliable Without Borrow";
                Worksheet.Cells[1, 6] = "USD Value";
                Worksheet.Cells[1, 7] = "Spot Borrow";

                for (int i = 0; i < analysis.coin.Length; i++)
                {
                    Worksheet.Cells[i + 2, 1] = analysis.coin[i];
                    Worksheet.Cells[i + 2, 2] = analysis.total[i];
                    Worksheet.Cells[i + 2, 3] = analysis.free[i];
                    Worksheet.Cells[i + 2, 4] = analysis.avaliableForWithdrawal[i];
                    Worksheet.Cells[i + 2, 5] = analysis.avaliableWithoutBorrow[i];
                    Worksheet.Cells[i + 2, 6] = analysis.usdValue[i];
                    Worksheet.Cells[i + 2, 7] = analysis.spotBorrow[i];

                }


            }

            else
            {
                Console.WriteLine("Historical Data Error");
            }

        }

        public void Create_Excel_Historical_Data(Historical_Price_Analysis analysis, Excel.Application Application, int start)
        {
            if (Application != null)
            {
                Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();


                Worksheet.Cells[1, 1] = "time";
                Worksheet.Cells[1, 2] = "open";
                Worksheet.Cells[1, 3] = "close";
                Worksheet.Cells[1, 4] = "high";
                Worksheet.Cells[1, 5] = "low";
                Worksheet.Cells[1, 6] = "volume";

                int NumberOfRow = -1;
                int length = analysis.high.Length;
                if (length - start > 5000)
                {
                    NumberOfRow = 5000;
                }
                else
                {
                    NumberOfRow = length - start;
                }


                for (int i = 0; i < NumberOfRow; i++)
                {
                    Worksheet.Cells[i + 2, 1] = analysis.time[start + i];
                    Worksheet.Cells[i + 2, 2] = analysis.open[start + i];
                    Worksheet.Cells[i + 2, 3] = analysis.close[start + i];
                    Worksheet.Cells[i + 2, 4] = analysis.high[start + i];
                    Worksheet.Cells[i + 2, 5] = analysis.low[start + i];
                    Worksheet.Cells[i + 2, 6] = analysis.volume[start + i];

                }
            }
            else
            {
                Console.WriteLine("Historical Data Error");
            }

        }



        // End of Historical Prices

        public async Task<dynamic> GetBalancesSnapshot(int request_id)
        {
            var path = $"api/historical_balances/requests/{request_id}";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/historical_balances/requests/{request_id}", "");

            var result = await CallAsyncSign(HttpMethod.Get, path, sign);

            return ParseResponce(result);
        }

        // Withdrawal History

        public class DepositHistory
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public int id { get; set; }
                public string coin { get; set; }
                public float? size { get; set; }
                public string status { get; set; }
                public DateTime time { get; set; }
                public DateTime? confirmedTime { get; set; }
                public object uploadedFile { get; set; }
                public object uploadedFileName { get; set; }
                public object cancelReason { get; set; }
                public bool fiat { get; set; }
                public bool ach { get; set; }
                public string type { get; set; }
                public object supportTicketId { get; set; }
            }

        }

        public class DepositHistoryList
        {

            public int[] id;
            public string[] coin;
            public float?[] size;
            public string[] status;
            public DateTime[] time;
            public DateTime?[] confirmedTime;
            public string[] type;

            public DepositHistoryList(int length)
            {
                id = new int[length];
                type = new string[length];
                coin = new string[length];
                size = new float?[length];
                status = new string[length];
                time = new DateTime[length];
                confirmedTime = new DateTime?[length];
            }

        }
        // Analysis

        public DepositHistoryList DepositHistoryAnalysis(string deposit_history)
        {

            DepositHistory.Rootobject history = JsonConvert.DeserializeObject<DepositHistory.Rootobject>(deposit_history);
            DepositHistoryList historylist = new DepositHistoryList(history.result.Length);

            for (int i = 0; i < history.result.Length; i++)
            {
                historylist.id[i] = history.result[i].id;
                historylist.type[i] = history.result[i].type;
                historylist.coin[i] = history.result[i].coin;
                historylist.size[i] = history.result[i].size;
                historylist.status[i] = history.result[i].status;
                historylist.confirmedTime[i] = history.result[i].confirmedTime;
                historylist.time[i] = history.result[i].time;
            }

            return historylist;

        }

        public void Create_DepositoryHistory_Worksheet(DepositHistoryList depositHistoryList, Excel.Application application)
        {

            if (application == null)
            {
                Console.WriteLine("Create_DepositoryHistory_Worksheet: Application is Null");
            }

            if (depositHistoryList != null)
            {

                Excel.Worksheet Worksheet = (Excel.Worksheet)application.Worksheets.Add();
                Worksheet.Cells[1, 1] = "ID";
                Worksheet.Cells[1, 2] = "Type";
                Worksheet.Cells[1, 3] = "Coin";
                Worksheet.Cells[1, 4] = "Size";
                Worksheet.Cells[1, 5] = "Status";
                Worksheet.Cells[1, 6] = "Time";
                Worksheet.Cells[1, 7] = "Confirmed Time";



                for (int i = 0; i < depositHistoryList.status.Length; i++)
                {
                    Worksheet.Cells[i + 2, 1] = depositHistoryList.id[i];
                    Worksheet.Cells[i + 2, 2] = depositHistoryList.type[i];
                    Worksheet.Cells[i + 2, 3] = depositHistoryList.coin[i];
                    Worksheet.Cells[i + 2, 4] = depositHistoryList.size[i];
                    Worksheet.Cells[i + 2, 5] = depositHistoryList.status[i];
                    Worksheet.Cells[i + 2, 6] = depositHistoryList.time[i];
                    Worksheet.Cells[i + 2, 7] = depositHistoryList.confirmedTime[i];
                }

            }
            else
            {
                Console.WriteLine("DepositHistoryList is null");
            }
        }

        //End of Withdrawal History

        // Withdrawal History

        public class Withdrawal
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public string coin { get; set; }
                public string address { get; set; }
                public object tag { get; set; }
                public int fee { get; set; }
                public int id { get; set; }
                public float size { get; set; }
                public string status { get; set; }
                public DateTime time { get; set; }
                public string method { get; set; }
                public string txid { get; set; }
            }
        }

        public class WithdrawalList
        {
            public string[] coin;
            public string[] address;
            public int[] fee;
            public int[] id;
            public float[] size;
            public string[] status;
            public DateTime[] time;
            public string[] method;
            public string[] txid;

            public WithdrawalList(int length)
            {
                coin = new string[length];
                address = new string[length];
                fee = new int[length];
                id = new int[length];
                size = new float[length];
                status = new string[length];
                time = new DateTime[length];
                method = new string[length];
                txid = new string[length];

            }
        }


        public WithdrawalList WithdrawalHistoryAnalysis(string withdrawal_list)
        {
            Withdrawal.Rootobject history = JsonConvert.DeserializeObject<Withdrawal.Rootobject>(withdrawal_list);
            WithdrawalList Withdrawal_list = new WithdrawalList(history.result.Length);
            for (int i = 0; i < Withdrawal_list.size.Length; i++)
            {
                Withdrawal_list.coin[i] = history.result[i].coin;
                Withdrawal_list.address[i] = history.result[i].address;
                Withdrawal_list.fee[i] = history.result[i].fee;
                Withdrawal_list.id[i] = history.result[i].id;
                Withdrawal_list.size[i] = history.result[i].size;
                Withdrawal_list.status[i] = history.result[i].status;
                Withdrawal_list.time[i] = history.result[i].time;
                Withdrawal_list.method[i] = history.result[i].method;
                Withdrawal_list.txid[i] = history.result[i].txid;
            }

            return Withdrawal_list;

        }

        public void Create_WithdrawalHistory_Worhsheet(WithdrawalList list, Excel.Application application)
        {

            if (application != null)
            {
                Excel.Worksheet Worksheet = (Excel.Worksheet)application.Worksheets.Add();

                Worksheet.Cells[1, 1] = "Date";
                Worksheet.Cells[1, 2] = "Coin";
                Worksheet.Cells[1, 3] = "Size";
                Worksheet.Cells[1, 4] = "Address";
                Worksheet.Cells[1, 5] = "fee";
                Worksheet.Cells[1, 6] = "status";
                Worksheet.Cells[1, 7] = "method";
                Worksheet.Cells[1, 8] = "id";
                Worksheet.Cells[1, 9] = "txid";

                for (int i = 0; i < list.status.Length; i++)
                {
                    Worksheet.Cells[i + 2, 1] = list.id[i];
                    Worksheet.Cells[i + 2, 2] = list.coin[i];
                    Worksheet.Cells[i + 2, 3] = list.address[i];
                    Worksheet.Cells[i + 2, 4] = list.fee[i];
                    Worksheet.Cells[i + 2, 5] = list.size[i];
                    Worksheet.Cells[i + 2, 6] = list.status[i];
                    Worksheet.Cells[i + 2, 7] = list.time[i];
                    Worksheet.Cells[i + 2, 8] = list.method[i];
                    Worksheet.Cells[i + 2, 9] = list.txid[i];

                }
            }

            {
                Console.WriteLine("Application (Withdrawal History is Null");
            }

        }
        // End of Withdrawal History

        public void Rename_Excel(Excel.Application application)
        {
            application.Worksheets[1].Name = "Trade History";
            application.Worksheets[2].Name = "Tdasds";
            application.Worksheets[3].Name = "Withdrawal History";
            application.Worksheets[4].Name = "Deposit History";
            application.Worksheets[5].Name = "Historical BTC-PERP Daily";
            application.Worksheets[6].Name = "Historical BTC-PERP Weeky";
            application.Worksheets[7].Name = "Historical ETH-PERP Daily";
            application.Worksheets[8].Name = "Historical ETH-PERP Weekly";
            application.Worksheets[9].Name = "BreakDown of Balance";
        }

        public void Rename_Excel2(Excel.Application application)
        {
            application.Worksheets[1].Name = "ETH-PERP (15mins)";
            application.Worksheets[2].Name = "BTC-PERP (15mins)";
            ;
        }

        #endregion

        #region Orders

        public async Task<dynamic> PlaceOrderAsync(string instrument, SideType side, decimal price, OrderType orderType, decimal amount, bool reduceOnly = false)
        {
            var path = $"api/orders";

            var body =
                $"{{\"market\": \"{instrument}\"," +
                $"\"side\": \"{side.ToString()}\"," +
                $"\"price\": {price}," +
                $"\"type\": \"{orderType.ToString()}\"," +
                $"\"size\": {amount}," +
                $"\"reduceOnly\": {reduceOnly.ToString().ToLower()}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/orders", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);

            return ParseResponce(result);
        }

        public async Task<dynamic> PlaceStopOrderAsync(string instrument, SideType side, decimal triggerPrice, decimal amount, bool reduceOnly = false)
        {
            var path = $"api/conditional_orders";

            var body =
                $"{{\"market\": \"{instrument}\"," +
                $"\"side\": \"{side.ToString()}\"," +
                $"\"triggerPrice\": {triggerPrice}," +
                $"\"type\": \"stop\"," +
                $"\"size\": {amount}," +
                $"\"reduceOnly\": {reduceOnly.ToString().ToLower()}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/conditional_orders", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);

            return ParseResponce(result);
        }

        public async Task<dynamic> PlaceTrailingStopOrderAsync(string instrument, SideType side, decimal trailValue, decimal amount, bool reduceOnly = false)
        {
            var path = $"api/conditional_orders";

            var body =
                $"{{\"market\": \"{instrument}\"," +
                $"\"side\": \"{side.ToString()}\"," +
                $"\"trailValue\": {trailValue}," +
                $"\"type\": \"trailingStop\"," +
                $"\"size\": {amount}," +
                $"\"reduceOnly\": {reduceOnly.ToString().ToLower()}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/conditional_orders", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);

            return ParseResponce(result);
        }

        public async Task<dynamic> PlaceTakeProfitOrderAsync(string instrument, SideType side, decimal triggerPrice, decimal amount, bool reduceOnly = false)
        {
            var path = $"api/conditional_orders";

            var body =
                $"{{\"market\": \"{instrument}\"," +
                $"\"side\": \"{side.ToString()}\"," +
                $"\"triggerPrice\": {triggerPrice}," +
                $"\"type\": \"takeProfit\"," +
                $"\"size\": {amount}," +
                $"\"reduceOnly\": {reduceOnly.ToString().ToLower()}}}";

            var sign = GenerateSignature(HttpMethod.Post, "/api/conditional_orders", body);
            var result = await CallAsyncSign(HttpMethod.Post, path, sign, body);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetOpenOrdersAsync(string instrument)
        {
            var path = $"api/orders?market={instrument}";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/orders?market={instrument}", "");

            var result = await CallAsyncSign(HttpMethod.Get, path, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetOrderStatusAsync(string id)
        {
            var resultString = $"api/orders/{id}";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/orders/{id}", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetOrderStatusByClientIdAsync(string clientOrderId)
        {
            var resultString = $"api/orders/by_client_id/{clientOrderId}";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/orders/by_client_id/{clientOrderId}", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> CancelOrderAsync(string id)
        {
            var resultString = $"api/orders/{id}";

            var sign = GenerateSignature(HttpMethod.Delete, $"/api/orders/{id}", "");

            var result = await CallAsyncSign(HttpMethod.Delete, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> CancelOrderByClientIdAsync(string clientOrderId)
        {
            var resultString = $"api/orders/by_client_id/{clientOrderId}";

            var sign = GenerateSignature(HttpMethod.Delete, $"/api/orders/by_client_id/{clientOrderId}", "");

            var result = await CallAsyncSign(HttpMethod.Delete, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> CancelAllOrdersAsync(string instrument)
        {
            var resultString = $"api/orders";

            var body =
                $"{{\"market\": \"{instrument}\"}}";

            var sign = GenerateSignature(HttpMethod.Delete, $"/api/orders", body);

            var result = await CallAsyncSign(HttpMethod.Delete, resultString, sign, body);

            return ParseResponce(result);
        }

        #endregion

        #region Fills

        public class Fills
        {
            public class Rootobject
            {
                public bool success { get; set; }
                public Result[] result { get; set; }
            }

            public class Result
            {
                public long id { get; set; }
                public string market { get; set; }
                public string future { get; set; }
                public string baseCurrency { get; set; }
                public string quoteCurrency { get; set; }
                public string type { get; set; }
                public string side { get; set; }
                public float price { get; set; }
                public float size { get; set; }
                public long orderId { get; set; }
                public DateTime time { get; set; }
                public long tradeId { get; set; }
                public float feeRate { get; set; }
                public float fee { get; set; }
                public string feeCurrency { get; set; }
                public string liquidity { get; set; }
            }
        }

        public class Fills_List
        {
            public long[] id { get; set; }
            public string[] market { get; set; }
            public string[] future { get; set; }
            public string[] baseCurrency { get; set; }
            public string[] quoteCurrency { get; set; }
            public string[] type { get; set; }
            public string[] side { get; set; }
            public float[] price { get; set; }
            public float[] size { get; set; }
            public long[] orderId { get; set; }
            public DateTime[] time { get; set; }
            public long[] tradeId { get; set; }
            public float[] feeRate { get; set; }
            public float[] fee { get; set; }
            public string[] feeCurrency { get; set; }
            public string[] liquidity { get; set; }

            public Fills_List(int size)
            {

                id = new long[size];
                market = new string[size];
                future = new string[size];
                baseCurrency = new string[size];
                quoteCurrency = new string[size];
                type = new string[size];
                side = new string[size];
                price = new float[size];
                this.size = new float[size];
                orderId = new long[size];
                time = new DateTime[size];
                tradeId = new long[size];
                feeRate = new float[size];
                fee = new float[size];
                feeCurrency = new string[size];
                liquidity = new string[size];
            }

        }

        public async Task<dynamic> GetFillsAsync(DateTime start, DateTime end)
        {
            var resultString = $"api/fills?start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var sign = GenerateSignature(HttpMethod.Get, $"/{resultString}", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public Fills_List GetFillsList(DateTime start, DateTime end)
        {

            Fills.Rootobject fills = JsonConvert.DeserializeObject<Fills.Rootobject>(GetFillsAsync(start, end).Result);
            Fills_List fillsList = new Fills_List(fills.result.Length);

            for (int i = 0; i < fills.result.Length; i++)
            {

                fillsList.id[i] = fills.result[i].id;
                fillsList.market[i] = fills.result[i].market;
                fillsList.future[i] = fills.result[i].future;
                fillsList.baseCurrency[i] = fills.result[i].baseCurrency;
                fillsList.quoteCurrency[i] = fills.result[i].quoteCurrency;
                fillsList.type[i] = fills.result[i].type;
                fillsList.side[i] = fills.result[i].side;
                fillsList.price[i] = fills.result[i].price;
                fillsList.size[i] = fills.result[i].size;
                fillsList.orderId[i] = fills.result[i].orderId;
                fillsList.time[i] = fills.result[i].time;
                fillsList.tradeId[i] = fills.result[i].tradeId;
                fillsList.feeRate[i] = fills.result[i].feeRate;
                fillsList.fee[i] = fills.result[i].fee;
                fillsList.feeCurrency[i] = fills.result[i].feeCurrency;
                fillsList.liquidity[i] = fills.result[i].liquidity;

            }
            return fillsList;

        }


        public void Create_Trade_History_Excel(Fills_List fills_list, Excel.Application Application)
        {
            if (Application != null)
            {

                Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();

                Worksheet.Cells[1, 1] = "Time";
                Worksheet.Cells[1, 2] = "Market";
                Worksheet.Cells[1, 3] = "Side";
                Worksheet.Cells[1, 4] = "Order Type";
                Worksheet.Cells[1, 5] = "Size";
                //Worksheet.Cells[1, 6] = "Price";
                Worksheet.Cells[1, 6] = "Total";
                Worksheet.Cells[1, 7] = "fee";
                Worksheet.Cells[1, 8] = "fee Currency";
                Worksheet.Cells[1, 9] = "fee Rate";
                Worksheet.Cells[1, 10] = "future";
                Worksheet.Cells[1, 11] = "Base Currency";
                Worksheet.Cells[1, 12] = "QuoteCurrency";
                Worksheet.Cells[1, 13] = "Liquidity";
                Worksheet.Cells[1, 14] = "id";
                Worksheet.Cells[1, 15] = "Trade Id";
                Worksheet.Cells[1, 16] = "Order Id";



                for (int i = 0; i < fills_list.fee.Length; i++)
                {
                    Worksheet.Cells[i + 2, 1] = fills_list.time[i];
                    Worksheet.Cells[i + 2, 2] = fills_list.market[i];
                    Worksheet.Cells[i + 2, 3] = fills_list.side[i];
                    Worksheet.Cells[i + 2, 4] = fills_list.type[i];
                    Worksheet.Cells[i + 2, 5] = fills_list.size[i];
                    Worksheet.Cells[i + 2, 6] = fills_list.price[i];
                    Worksheet.Cells[i + 2, 7] = fills_list.fee[i];
                    Worksheet.Cells[i + 2, 8] = fills_list.feeCurrency[i];
                    Worksheet.Cells[i + 2, 9] = fills_list.feeRate[i];
                    Worksheet.Cells[i + 2, 10] = fills_list.future[i];
                    Worksheet.Cells[i + 2, 11] = fills_list.baseCurrency[i];
                    Worksheet.Cells[i + 2, 12] = fills_list.quoteCurrency[i];
                    Worksheet.Cells[i + 2, 13] = fills_list.liquidity[i];
                    Worksheet.Cells[i + 2, 14] = fills_list.id[i];
                    Worksheet.Cells[i + 2, 15] = fills_list.tradeId[i];
                    Worksheet.Cells[i + 2, 16] = fills_list.orderId[i];

                }
            }
            else
            {
                Console.WriteLine("Historical Data Error");
            }

        }



        #endregion

        #region Funding

        public async Task<dynamic> GetFundingPaymentAsync(DateTime start, DateTime end)
        {
            var resultString = $"api/funding_payments?start_time={Util.Util.GetSecondsFromEpochStart(start)}&end_time={Util.Util.GetSecondsFromEpochStart(end)}";

            var sign = GenerateSignature(HttpMethod.Get, $"/{resultString}", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        #endregion

        #region Leveraged Tokens

        public async Task<dynamic> GetLeveragedTokensListAsync()
        {
            var resultString = $"api/lt/tokens";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetTokenInfoAsync(string tokenName)
        {
            var resultString = $"api/lt/{tokenName}";

            var result = await CallAsync(HttpMethod.Get, resultString);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetLeveragedTokenBalancesAsync()
        {
            var resultString = $"api/lt/balances";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/lt/balances", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetLeveragedTokenCreationListAsync()
        {
            var resultString = $"api/lt/creations";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/lt/creations", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> RequestLeveragedTokenCreationAsync(string tokenName, decimal size)
        {
            var resultString = $"api/lt/{tokenName}/create";

            var body = $"{{\"size\": {size}}}";

            var sign = GenerateSignature(HttpMethod.Post, $"/api/lt/{tokenName}/create", body);

            var result = await CallAsyncSign(HttpMethod.Post, resultString, sign, body);

            return ParseResponce(result);
        }

        public async Task<dynamic> GetLeveragedTokenRedemptionListAsync()
        {
            var resultString = $"api/lt/redemptions";

            var sign = GenerateSignature(HttpMethod.Get, $"/api/lt/redemptions", "");

            var result = await CallAsyncSign(HttpMethod.Get, resultString, sign);

            return ParseResponce(result);
        }

        public async Task<dynamic> RequestLeveragedTokenRedemptionAsync(string tokenName, decimal size)
        {
            var resultString = $"api/lt/{tokenName}/redeem";

            var body = $"{{\"size\": {size}}}";

            var sign = GenerateSignature(HttpMethod.Post, $"/api/lt/{tokenName}/redeem", body);

            var result = await CallAsyncSign(HttpMethod.Post, resultString, sign, body);

            return ParseResponce(result);
        }

        #endregion

        #region Util

        // Testnet

        private async Task<string> CallAsync(HttpMethod method, string endpoint, string body = null)
        {
            var request = new HttpRequestMessage(method, endpoint);

            if (body != null)
            {
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");
            }

            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);

            var result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            return result;
        }


        private async Task<string> CallAsyncSign(HttpMethod method, string endpoint, string sign, string body = null)
        {
            var request = new HttpRequestMessage(method, endpoint);

            if (body != null)
            {
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");
            }

            request.Headers.Add("FTX-KEY", _client.ApiKey);
            request.Headers.Add("FTX-SIGN", sign);
            request.Headers.Add("FTX-TS", _nonce.ToString());

            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);

            var result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            return result;
        }

        private string GenerateSignature(HttpMethod method, string url, string requestBody)
        {
            _nonce = GetNonce();
            var signature = $"{_nonce}{method.ToString().ToUpper()}{url}{requestBody}";
            var hash = _hashMaker.ComputeHash(Encoding.UTF8.GetBytes(signature));
            var hashStringBase64 = BitConverter.ToString(hash).Replace("-", string.Empty);
            return hashStringBase64.ToLower();
        }

        private long GetNonce()
        {
            return Util.Util.GetMillisecondsFromEpochStart();
        }

        private dynamic ParseResponce(string responce)
        {
            return (dynamic)responce;
        }


        #endregion


        #region Helper_Function

        public class Account_record
        {

            public class Rootobject
            {
                public bool success { get; set; }
                public Result result { get; set; }
            }

            public class Result
            {
                public int id { get; set; }
                public string[] accounts { get; set; }
                public string time { get; set; }
                public string endTime { get; set; }
                public string status { get; set; }
                public bool error { get; set; }
                public Result1[] results { get; set; }
            }

            public class Result1
            {
                public string account { get; set; }
                public string ticker { get; set; }
                public float size { get; set; }
                public float price { get; set; }
            }

        }

        public void GetAccountValue(string[] accounts, Excel.Application Application, int[] snapshot)
        {

            int endTime = Convert.ToInt32((DateTimeOffset.UtcNow.ToUnixTimeSeconds() / 86400) * 86400);
            int startTime = 1645977600; // 28 Feb
            int NumberOfDays = Convert.ToInt32((endTime - startTime) / 86400);


            Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();
            Worksheet.Cells[1, 1] = "Date";
            Worksheet.Cells[1, 2] = "Value";


            for (int i = 0; i < NumberOfDays - 1; i++)
            {
                Task.Delay(1000).Wait();
                Account_record.Rootobject temp = JsonConvert.DeserializeObject<Account_record.Rootobject>(GetHistoricalBalancesWith_ID(snapshot[i]).Result);
                float value = 0;
                for (int j = 0; j < temp.result.results.Length; j++)
                {
                    if (temp.result.results[j].ticker == "USD")
                    {
                        value = value + temp.result.results[j].size;
                    }
                    else
                    {
                        value = value + temp.result.results[j].price * temp.result.results[j].size;
                    }
                    Console.WriteLine(value);
                }
                Console.WriteLine("SubTotal = " + value);
                Worksheet.Cells[i + 2, 2] = value;
                Worksheet.Cells[i + 2, 1] = temp.result.endTime;
            }
        }
        public int[] GetTimeStamp()
        {

            long Date_Today_TimeStamp_zero = (DateTimeOffset.UtcNow.ToUnixTimeSeconds() / 86400) * 86400;
            int Start_date_TimeStamp = 1645977600; // 28 Feb
            int NumberOfWeeks = Convert.ToInt32((Date_Today_TimeStamp_zero - 1645977600) / 86400);


            int[] each_day_array = new int[NumberOfWeeks];

            for (int i = 0; i < NumberOfWeeks; i++)
            {
                each_day_array[i] = Start_date_TimeStamp + i * 86400;
            }

            return each_day_array;
        }

        public int[] GetSnapshotList(int[] TimeStamp, string[] accountList)
        {
            int[] SnapshotList = new int[TimeStamp.Length];
            for (int i = 0; i < TimeStamp.Length; i++)
            {
                int snap_temp = 0;
                while (snap_temp == 0)
                {
                    Random rnd = new Random();
                    Console.WriteLine(i);
                    Task.Delay(rnd.Next(1000)).Wait();
                    snap_temp = Get_SnapPoint(accountList, TimeStamp[i]);
                    SnapshotList[i] = snap_temp;

                }
            }

            Console.WriteLine("Done");
            return SnapshotList;

        }

        #endregion


        #region Bar_Data

        public void Create_Excel_Bar_Data(string market, int timeframe, DateTime dtDateStart, DateTime dtDateEnd, Excel.Application Application)
        {

            Excel.Worksheet Worksheet = (Excel.Worksheet)Application.Sheets.Add();

            var dateStart = dtDateStart.AddSeconds(0).ToLocalTime();
            var dateEnd = dtDateEnd.AddSeconds(0).ToLocalTime();

            int unixStartTime = 1559881511;
            long temp = ((DateTimeOffset)dateEnd).ToUnixTimeSeconds();
            int unixEndTime = Convert.ToInt32(temp);

            int page = ((unixEndTime - unixStartTime) / timeframe / 5000) + 1;

            int counter = 0;

            for (int j = 0; j < page; j++)
            {
                var dateEnd_loop = dateStart.AddSeconds(timeframe * 5000 * (j + 1)).ToLocalTime();
                var temp2 = ((DateTimeOffset)dateEnd_loop).ToUnixTimeSeconds();

                Console.WriteLine(dateStart.AddSeconds(timeframe * 5000 * (j)).ToLocalTime());
                Console.WriteLine("The theoretical ending TimeStamp = " + temp2);
                Console.WriteLine(" The time frame of current time = " + unixEndTime);

                BarData.Rootobject records = new BarData.Rootobject();
                if (unixEndTime < temp2)
                {

                    records = JsonConvert.DeserializeObject<BarData.Rootobject>(GetHistoricalPricesAsync(market, timeframe, 100000, dateStart.AddSeconds(5000 * timeframe * j).ToLocalTime(), dtDateEnd).Result);
                }
                else
                {
                    records = JsonConvert.DeserializeObject<BarData.Rootobject>(GetHistoricalPricesAsync(market, timeframe, 100000, dateStart.AddSeconds(5000 * timeframe * j).ToLocalTime(), dateStart.AddSeconds(timeframe * 5000 * (j + 1)).ToLocalTime()).Result);

                }

                int length = records.result.Length;

                Historical_Price_Analysis result = new Historical_Price_Analysis(length);

                for (int i = 0; i < records.result.Length; i++)
                {

                    Worksheet.Cells[counter + 2 + i, 1] = records.result[i].startTime;
                    Worksheet.Cells[counter + 2 + i, 2] = records.result[i].open;
                    Worksheet.Cells[counter + 2 + i, 3] = records.result[i].high;
                    Worksheet.Cells[counter + 2 + i, 4] = records.result[i].low;
                    Worksheet.Cells[counter + 2 + i, 5] = records.result[i].close;
                    Worksheet.Cells[counter + 2 + i, 6] = records.result[i].volume;
                }

                counter = counter + length;


            }

            Console.WriteLine("The Expected number of results = ", counter);


        }
    }

    #endregion
}
