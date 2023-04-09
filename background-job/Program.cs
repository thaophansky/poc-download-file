using Hangfire;

namespace background_job;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.IO;
class Program
{
    static void Main(string[] args)
    {
	    const string jobId = "craw-test-rd";
	    GlobalConfiguration.Configuration
		    .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
		    .UseColouredConsoleLogProvider()
		    .UseSimpleAssemblyNameTypeSerializer()
		    .UseRecommendedSerializerSettings()
		    .UseInMemoryStorage();
        RecurringJob.AddOrUpdate(jobId, () => Test(), Cron.Hourly);
        RecurringJob.TriggerJob(jobId);
	    using (var server = new BackgroundJobServer())
	    {
		    Console.ReadLine();
	    }
    }
    static void Test()
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

        var spreadsheetId = "1X2olReH_7pZy29wtne36kI8mpFk4YxJLaJPezMWxPlw";
        var range = "item!A1:P3000";
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

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
