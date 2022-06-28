using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace CatchTest
{
    class MainClass
    {
        public static async Task Main(string[] args)
        {
            //初始化数据
            var serviceProvider = new ServiceCollection().AddHttpClient().BuildServiceProvider();
            var httpClientFactory = serviceProvider.GetService<IHttpClientFactory>();

            var ids = GetKeyValuePairs();
            var dates = GetDate();
            var uri = "http://yun.lwbsq.com/data/history";
            var setting = new JsonSerializerSettings
            {
                ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver(),
            };

            foreach (var item in ids)
            {
                Dictionary<string, List<DataClass>> keyValuePairs = new Dictionary<string, List<DataClass>>();
                for (int i = 0; i < dates.Count-1; i++)
                {

                    var client = httpClientFactory.CreateClient();
                    client.DefaultRequestHeaders.Add("Accept", "application/json, text/javascript, */*; q=0.01");
                    client.DefaultRequestHeaders.Add("Cookie", "JSESSIONID=4F5E0CE0BC6704134843E66BE73147A8; SECKEY_ABVK=DPFhaTdMK571HhRsyLFejYOB15bd4TEal1Fd7hJxikE%3D; BMAP_SECKEY=DPFhaTdMK571HhRsyLFejapz75-1DG8DokWTZ5fkaFz3yIb8KaHL_KE9zR_ekzt8OHJfEoi80ZkIscBioJsvJOjImlDWd7-lLmL_iC03BEAdAZFoP4pLDwNxSG1z6ji01PiYs5gyJm7W7FXCo4gGlJZ-rNq9kPBQd7DJCjAZLjN5C4r6VFUcLfSG-M8wLONG");
                    TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
                    var timestamp = Convert.ToInt64(ts.TotalMilliseconds);
                    var paramList = new List<KeyValuePair<string, string>>()
                    {
                        new KeyValuePair<string, string>("termId",item.Key),
                        new KeyValuePair<string, string>("beginTime",dates[i]),
                        new KeyValuePair<string, string>("endTime",dates[i+1]),
                        new KeyValuePair<string, string>("&_", $"{timestamp}"),
                    };
                    var result = await client.GetAsync($"{uri}?{string.Join("&", paramList.Select(p => System.Net.WebUtility.UrlEncode(p.Key) + "=" + System.Net.WebUtility.UrlEncode(p.Value)))}");
                    if (result.IsSuccessStatusCode)
                    {
                        var responseStream = await result.Content.ReadAsStringAsync();
                        try
                        {
                            var dy = JsonConvert.DeserializeObject<res>(responseStream);
                            //var dy = System.Text.Json.JsonSerializer.Deserialize<res>(responseStream);
                            if (dy.code == 1000)
                            {
                                var res = dy.data;
                                var records = new List<DataClass>();
                                foreach (var re in res)
                                {
                                    System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));//当地时区
                                    records.Add(new DataClass { termname = re.termname, temp = re.temp, humidity = re.humidity, longitude = re.longitude, latitude = re.latitude, time = startTime.AddMilliseconds(re.htime) });
                                }
                                keyValuePairs.Add(dates[i].Substring(0, 10), records);
                                Console.WriteLine($"完成 {item.Value} 时间段 {dates[i].Substring(0, 10)}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(responseStream);
                            i--;
                            Thread.Sleep(60000);
                            continue;
                        }
                        
                    }
                    else
                    {
                        Console.WriteLine(result.StatusCode);
                    }
                }
                Export(item.Value, keyValuePairs);
            }

            

        }

        private static Dictionary<string, string> GetKeyValuePairs()
        {

            Dictionary<string, string> keys = new Dictionary<string, string>();
            keys.Add("672585", "40049572-0-5温湿");
            keys.Add("672586", "40049572-0-5电导率");
            keys.Add("672587", "40049572-0-20温湿度");
            keys.Add("672588", "40049572-0-20电导率");
            keys.Add("672589", "40049572-0-35温湿度");
            keys.Add("672590", "40049572-0-35电导率");
            keys.Add("672591", "40049572-0-50温湿度");
            keys.Add("672592", "40049572-0-50电导率");
            keys.Add("672593", "40049572-0-65温湿度");
            keys.Add("672594", "40049572-0-65电导率");
            keys.Add("672595", "40049572-7-65温湿度");
            keys.Add("672596", "40049572-7-65电导率");
            keys.Add("672597", "40049572-7-50温湿度");
            keys.Add("672598", "40049572-7-50电导率");
            keys.Add("672599", "40049572-7-35温湿度");
            keys.Add("672600", "40049572-7-35电导率");
            keys.Add("672601", "40049572-7-20温湿度");
            keys.Add("672602", "40049572-7-20电导率");
            keys.Add("672603", "40049572-7-5温湿度");
            keys.Add("672604", "40049572-7-5电导率");
            keys.Add("672605", "40049572-14-5温湿度");
            keys.Add("672606", "40049572-14-5电导率");
            keys.Add("672607", "40049572-14-20温湿度");
            keys.Add("672608", "40049572-14-20电导率");
            keys.Add("672609", "40049572-14-35温湿度");
            keys.Add("672610", "40049572-14-35电导率");
            keys.Add("672611", "40049572-14-50温湿度");
            keys.Add("672612", "40049572-14-50电导率");
            keys.Add("672613", "40049572-14-65温湿度");
            keys.Add("672614", "40049572-14-65电导率");
            return keys;
        }

        private static List<string> GetDate()
        {
            List<string> vs = new List<string>()
            {
                "2021-06-01 00:00",
                "2021-07-01 00:00",
                "2021-08-01 00:00",
                "2021-09-01 00:00",
                "2021-10-01 00:00",
                "2021-11-01 00:00",
            };
            return vs;
        }

        private static void Export(string name, Dictionary<string, List<DataClass>> keyValuePairs)
        {
            #region 导出
            //创建工作簿
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            //创建Sheet页
            #region MyRegion
            ICellStyle style = hssfworkbook.CreateCellStyle();
            //设置边框
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            //设置单元格的样式：水平对齐居中
            style.Alignment = HorizontalAlignment.Left;
            //新建一个字体样式对象
            IFont font = hssfworkbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            //使用SetFont方法将字体样式添加到单元格样式中 
            style.SetFont(font);

            #endregion

            #region MyRegion
            ICellStyle Istyle2 = hssfworkbook.CreateCellStyle();
            //设置边框
            Istyle2.BorderTop = BorderStyle.Thin;
            Istyle2.BorderBottom = BorderStyle.Thin;
            Istyle2.BorderLeft = BorderStyle.Thin;
            Istyle2.BorderRight = BorderStyle.Thin;
            //设置单元格的样式：水平对齐居中
            Istyle2.Alignment = HorizontalAlignment.Center;
            //新建一个字体样式对象
            IFont Ifont2 = hssfworkbook.CreateFont();
            Ifont2.FontName = "宋体";
            Ifont2.FontHeightInPoints = 11;
            //设置字体加粗样式
            Ifont2.Boldweight = (short)FontBoldWeight.Bold;
            //使用SetFont方法将字体样式添加到单元格样式中 
            Istyle2.SetFont(Ifont2);
            #endregion
            try
            {
                foreach (var item in keyValuePairs)
                {
                    //创建Sheet页
                    ISheet sheet = hssfworkbook.CreateSheet(item.Key);
                    int RowIndex = 1;

                    #region  如果为第一行
                    IRow IRow = sheet.CreateRow(0);
                    for (int j = 0; j < typeof(Data).GetProperties().Length; j++)
                    {
                        ICell Icell2 = IRow.CreateCell(j);
                        
                        //将新的样式赋给单元格
                        Icell2.CellStyle = Istyle2;
                        Icell2.SetCellValue(typeof(Data).GetProperties()[j].Name);
                    }
                    #endregion
                    var sheetvalue = item.Value;
                    for (int k = 0; k < sheetvalue.Count; k++)
                    {
                        IRow row = sheet.CreateRow(RowIndex);
                        int l = 0;
                        foreach (var prop in typeof(Data).GetProperties())
                        {
                            ICell cell = row.CreateCell(l);
                            //将新的样式赋给单元格
                            cell.CellStyle = style;
                            cell.SetCellValue(Convert.ToString(prop.GetValue(sheetvalue[k], null)));

                            l++;
                        }
                        RowIndex++;
                    }
                }

            }
            catch (Exception ex) { }
            string Path = @"/Users/houjian/Desktop/exports";
            if (!System.IO.Directory.Exists(Path))
                System.IO.Directory.CreateDirectory(Path);
            string fileName = name + ".xls";
            using (FileStream file = new FileStream(System.IO.Path.Combine(Path,fileName), FileMode.Create))
            {
                hssfworkbook.Write(file);  //创建test.xls文件。
                file.Close();
            }
            Console.WriteLine($"导出完成{fileName}");
            #endregion
        }
    }

    #region class

    public class res {

        public int code { get; set; }
        public string message { get; set; }
        public List<Data> data { get; set; }
    }

    public struct Data
    {
        public string termname { get; set; }
        public decimal temp { get; set; }
        public decimal humidity { get; set; }
        public double longitude { get; set; }
        public double latitude { get; set; }
        public long htime { get; set; }
        public DateTime time { get; set; }
    }

    public class DataClass
    {
        public string termname { get; set; }
        public decimal temp { get; set; }
        public decimal humidity { get; set; }
        public double longitude { get; set; }
        public double latitude { get; set; }
        public long htime { get; set; }
        public DateTime time { get; set; }
    }

    #endregion
}