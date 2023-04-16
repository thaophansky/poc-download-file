using Hangfire;
using Hangfire.MemoryStorage;

namespace background_job;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Configuration;
using System.IO;
class Program
{
    static void Main(string[] args)
    {
        var builder = new ConfigurationBuilder()
                   .AddJsonFile($"appsettings.json", true, true);
        var config = builder.Build();
        var excelFileId = config["HeoConfig:ExcelFileId"];
        Console.WriteLine(excelFileId);
       
        const string jobId = "craw-test-rd";
        GlobalConfiguration.Configuration
            .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
            .UseColouredConsoleLogProvider()
            .UseSimpleAssemblyNameTypeSerializer()
            .UseRecommendedSerializerSettings()
            .UseMemoryStorage();
        RecurringJob.AddOrUpdate(jobId, () => Test(excelFileId), Cron.Minutely);
        RecurringJob.TriggerJob(jobId);
        using (var server = new BackgroundJobServer())
        {
            Console.ReadLine();
        }
    }
    public static void Test(string excelFileId)
    {
        UserCredential credential;

        using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        {
            string credPath = "token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                new[] { SheetsService.Scope.SpreadsheetsReadonly },
                "user",
                System.Threading.CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
        }

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "background-job",
        });

        var range = "item!A1:P3000";
        var request = service.Spreadsheets.Values.Get(excelFileId, range);

        var response = request.Execute();
        var values = response.Values;

        if (values != null && values.Count > 0)
        {
            using (var writer = new StreamWriter("output.csv"))
            {
                foreach (var row in values)
                {
                    writer.WriteLine(string.Join(",", row));
                }
            }
        }
        else
        {
            Console.WriteLine("No data found.");
        }
    }
}
