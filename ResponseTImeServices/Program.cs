using System.Diagnostics;
using System.Text;
using ResponseTImeServices;
using Newtonsoft.Json;
using ClosedXML.Excel;

Console.WriteLine("Start Testing!");

const string fileName = "Endpoint.json";
var currentDirectory = Directory.GetCurrentDirectory();
var filePath = Path.Combine(currentDirectory, fileName);

if (File.Exists(filePath))
{
    var jsonString = File.ReadAllText(filePath);
    
    var endpoints = JsonConvert.DeserializeObject<List<Endpoint>>(jsonString);
    var results = new List<ResponseTimeResult>();

    if (endpoints != null)
        foreach (var endpoint in endpoints)
        {
            var responseTime = await MeasureResponseTime(endpoint);
            results.Add(new ResponseTimeResult
            {
                Url = endpoint.Url,
                Method = endpoint.Method,
                ResponseTime = responseTime
            });

            Console.WriteLine($"URL: {endpoint.Url}, Method: {endpoint.Method}, Response Time: {responseTime} ms");
        }

    SaveToExcel(results);
}

Console.ReadKey();
return;

static async Task<long> MeasureResponseTime(Endpoint endpoint)
{
    using var client = new HttpClient();
    var request = new HttpRequestMessage
    {
        Method = new HttpMethod(endpoint.Method),
        RequestUri = new Uri(endpoint.Url)
    };

    if (endpoint.NeedParam && !string.IsNullOrEmpty(endpoint.Body))
    {
        request.Content = new StringContent(endpoint.Body, Encoding.UTF8, "application/json");
    }

    var stopwatch = Stopwatch.StartNew();
    using var response = await client.SendAsync(request);
    stopwatch.Stop();

    return stopwatch.ElapsedMilliseconds;
}

static void SaveToExcel(List<ResponseTimeResult> results)
{
    using var workbook = new XLWorkbook();
    var worksheet = workbook.Worksheets.Add("Response Times");
    worksheet.Cell(1, 1).Value = "URL";
    worksheet.Cell(1, 2).Value = "Method";
    worksheet.Cell(1, 3).Value = "Response Time (ms)";

    for (var i = 0; i < results.Count; i++)
    {
        worksheet.Cell(i + 2, 1).Value = results[i].Url;
        worksheet.Cell(i + 2, 2).Value = results[i].Method;
        worksheet.Cell(i + 2, 3).Value = results[i].ResponseTime;
    }

    var dateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
    workbook.SaveAs($"ResponseTimes{dateTime}.xlsx");
    Console.WriteLine("Results saved to ResponseTimes.xlsx");
}