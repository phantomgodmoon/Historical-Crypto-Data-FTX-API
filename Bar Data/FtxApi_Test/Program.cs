using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using FtxApi;
using Microsoft.Extensions.Configuration.Json;
using System.Web.Script.Serialization;
using System.Text.Json;
using FtxApi.Enums;
using Newtonsoft.Json;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;


namespace FtxApi_Test
{
    class Program
    {
        // Self-written Part
        static void Main()
        {
            // Input your Information
            string public_key = "dyIezmUoLDi_f41EC8VTLeZf7ewujMz1i-kZx6nE";
            string private_key = "lmSnTzQ9datCQtSR-NsiwesjPPvNuD8CUGk24XcU";
            string save_location2 = @"C:\Users\legend\Desktop\Hisotrical Data_BTC_Weekly\";


            // Log in with API Key and Secret Key and connect to the FTX API
            var client = new Client(public_key, private_key);
            var api = new FtxRestApi(client);
            var wsApi = new FtxWebSocketApi("wss://ftx.com/ws/");

            // Define the time

            DateTime dtDateTime = new DateTime(2019, 7, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            DateTime dtDateEnd = new DateTime(2022, 7, 29, 8, 0, 0, 0);

            var dateStart = dtDateTime.AddSeconds(0).ToUniversalTime();
            var dateEnd = dtDateEnd.AddSeconds(0).ToUniversalTime();

            // time frame = 15, 60 ,300, 900, 3600, 14400, 86400
            int timeframe = 86400*7;
            string BaseCurrency = "/BTC";
            string save_base = "";
            string time;
            if (timeframe == 86400)
            {
                time = "_1day";
            }
            else { time = "_1week"; }
            if (BaseCurrency == "/USD")
            {
                save_base = "USD";
            }
            if (BaseCurrency == "/BTC")
            {
                save_base = "BTC";
            }

            int save_base_length = save_base.Length;

            int unixStartTime = 1559881511; // Jun 7 2019
            long temp = ((DateTimeOffset)dateEnd).ToUnixTimeSeconds();
            int unixEndTime = Convert.ToInt32(temp);

            int page = ((unixEndTime - unixStartTime) / timeframe / 5000) + 1;

            /*string[] coin_list = { "ETH", "BTC", "SOL", "MATIC", "XRP", "BNB", "LTC", "AVAX", "FTT", "UNI", "BCH", "ATOM",
                "LINK", "APE", "DOT", "DOGE", "SNX","WBTC","AAVE", "CEL", "FTM", "SAND", "TRX", "DYDX",
                "LDO","GALA", "CRV", "YFI", "RAY", "SRM", "LOOKS", "AXS", "MANA","FXS" , "IMX", "GST","SHIB",
                "WAVES", "NEAR", "OMG", "KNC", "GRT", "POLIS","YFII", "STG", "RSR", "ENS", "ATLAS", "BOBA",
                "PAXG", "QI", "PEOPLE","STMX","1INCH", "TONCOIN","ALPHA", "COMP", "PERP", "ALGO", "TRYB","MKR",
                "CREAM", "CHR", "CVX", "CHZ", "CRO", "BIT","HOLY","AMPL", "MOB","BADGER","TRU", "HT", "AUDIO",
                "ENJ", "C98", "STEP","OXY", "VGX","BNT","GAL", "SPELL", "REEF", "BTT", "COMP" ,"AUDIO","SLP","ALICE",
                "YGG","CTX","PSG","NEXO","OKB","HNT","DENT","GODS","LINA","REN","GARI","LEO","JOE","ALCX","FIDA","ASD","SXP","HXRO","PROM",
                "STEP","TRU","INTER","ROOK","MER","WRX","GOG","MCB","MAPS","MNGO","COPE","AKRO","EDEN","BLT","JST","PORT","FRONT","GALFAN","KBTT","MTA",
                "INDI","MBS","CITY","TULIP","PUNDIX","PSY","SECO","BAR","DAWN","SLRS","MATH","HMT","UBXT","SOS","STARS","SNY","PRISM","GT","MTL",
                "PTU","ORBS","HUM","HGET","LUA"};*/

            List<string> coin_list = new List<string>();
            Rootobject records = JsonConvert.DeserializeObject<Rootobject>(api.GetMarketsAsync().Result);
            for (int i = 0; i < records.result.Length; i++) {
                if (records.result[i].name.Substring(records.result[i].name.Length - save_base_length)==save_base){
                    coin_list.Add(records.result[i].name);
                    Console.WriteLine(records.result[i].name+"\n");
                }
            }


            for (int i = 0; i < coin_list.Count; i++)
            {
                if (save_base == coin_list[i]) { continue; }
                string check_location = save_location2 + coin_list[i].Substring(0, coin_list[i].Length - save_base_length - 1) + save_base + time+".csv";
                if (!File.Exists(check_location))
                {
                    Console.WriteLine("Extracting data for " + coin_list[i]);
                    // Create Excel 
                    Excel.Application Application = api.create_Empty_Excel_Application();

                    for (int j = 0; j < page; j++)
                    {
                        int counter = 0;

                        var dateEnd_loop = dateStart.AddSeconds(timeframe * 5000 * (j + 1)).ToUniversalTime();
                        var temp2 = ((DateTimeOffset)dateEnd_loop).ToUnixTimeSeconds();

                        Console.WriteLine(dateStart.AddSeconds(timeframe * 5000 * (j)).ToUniversalTime());
                        Console.WriteLine(temp2);

                        if (timeframe == 60 * 60 * 24 * 7)
                        {
                            api.Create_Excel_Historical_Data(api.RecordOfHistoricalData_Spot(coin_list[i], timeframe, dateStart.AddSeconds(5000 * timeframe * j).ToUniversalTime(), dtDateEnd), Application, 0);
                        }
                        else if (unixEndTime < temp2)
                        {
                            api.Create_Excel_Historical_Data(api.RecordOfHistoricalData_Spot(coin_list[i], timeframe, dateStart.AddSeconds(5000 * timeframe * j).ToUniversalTime(), dtDateEnd), Application, counter);
                        }
                        else
                        {
                            api.Create_Excel_Historical_Data(api.RecordOfHistoricalData_Spot(coin_list[i], timeframe, dateStart.AddSeconds(5000 * timeframe * j).ToUniversalTime(), dateStart.AddSeconds(timeframe * 5000 * (j + 1)).ToUniversalTime()), Application, counter);
                        }

                    }

                    api.Save_Excel_File(Application, check_location);

                    Console.WriteLine("Done: data for " + coin_list[i]);
                }
                else
                {
                    Console.WriteLine(coin_list[i] + "Path Already exist");
                }
            }
        }

        private static async Task WebSocketTest(FtxWebSocketApi wsApi, Client client)
        {


            wsApi.OnWebSocketConnect += () =>
            {
                wsApi.SendCommand(FtxWebSockerRequestGenerator.GetAuthRequest(client));
                wsApi.SendCommand(FtxWebSockerRequestGenerator.GetSubscribeRequest("fills"));
                wsApi.SendCommand(FtxWebSockerRequestGenerator.GetSubscribeRequest("orders"));
            };

            await wsApi.Connect();
        }


        public class Rootobject
        {
            public bool success { get; set; }
            public Result[] result { get; set; }
        }

        public class Result
        {
            public string name { get; set; }
            public bool enabled { get; set; }
            public bool postOnly { get; set; }
            public float priceIncrement { get; set; }
            public float sizeIncrement { get; set; }
            public float minProvideSize { get; set; }
            public float last { get; set; }
            public float? bid { get; set; }
            public float? ask { get; set; }
            public float? price { get; set; }
            public string type { get; set; }
            public string futureType { get; set; }
            public string baseCurrency { get; set; }
            public bool isEtfMarket { get; set; }
            public string quoteCurrency { get; set; }
            public string underlying { get; set; }
            public bool restricted { get; set; }
            public bool highLeverageFeeExempt { get; set; }
            public float largeOrderThreshold { get; set; }
            public float change1h { get; set; }
            public float change24h { get; set; }
            public float changeBod { get; set; }
            public float quoteVolume24h { get; set; }
            public float volumeUsd24h { get; set; }
            public float priceHigh24h { get; set; }
            public float priceLow24h { get; set; }
            public bool tokenizedEquity { get; set; }
        }



    }



}