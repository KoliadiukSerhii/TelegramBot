using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace TelegramBotGmailCalendar.APIHelpers
{
    public static class CalendarHelpers
    {
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Telegram Bot Gmail Calendar Application";

        public static CalendarService GetCalendarService()
        {
            UserCredential credential;
            using (FileStream stream = new FileStream(Convert.ToString(ConfigurationManager.AppSettings["ClientInfo"]),
                FileMode.Open, FileAccess.Read))
            {
                String FolderPath = Convert.ToString(ConfigurationManager.AppSettings["CredentialsInfo"]);
                String FilePath = Path.Combine(FolderPath, "APITokenCredentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(FilePath, true)).Result;
            }

            CalendarService service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static EventsResource.ListRequest GetRequest(CalendarService service)
        {
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 5;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            return request;
        }
    }
}
