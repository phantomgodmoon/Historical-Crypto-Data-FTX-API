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
            string save_location2 = @"C:\Users\legend\Desktop\Hisotrical Data_USD_15mins\";

            // Log in with API Key and Secret Key and connect to the FTX API
            var client = new Client(public_key, private_key);
            var api = new FtxRestApi(client);
            var wsApi = new FtxWebSocketApi("wss://ftx.com/ws/");

            DateTime dtDateTime = new DateTime(2019, 7, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            DateTime dtDateEnd = new DateTime(2022, 7, 29, 8, 0, 0, 0);

            var dateStart = dtDateTime.AddSeconds(0).ToUniversalTime();
            var dateEnd = dtDateEnd.AddSeconds(0).ToUniversalTime();

            // time frame = 900, 3600, 14400, 86400
            int timeframe = 3600;
            string BaseCurrency = "/USD";
            string save_base = "";
            string time;
            if (timeframe == 86400)
            {
                time = "_1day";
            }
            else if (timeframe == 86400 * 7)
            {
                time = "_1week";
            }
            else if (timeframe == 3600)
            {
                time = "_1hour";
            }
            else if (timeframe == 900)
            {
                time = "_15mins";
            }
            else time = "ERROR";
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