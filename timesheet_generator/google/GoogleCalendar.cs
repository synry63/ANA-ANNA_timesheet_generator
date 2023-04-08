using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using timesheet_generator.resources;

namespace timesheet_generator.google
{
    public class GoogleCalendar
    {
        private readonly CalendarService _service;

        public GoogleCalendar(string credentialFilePath)
        {

            UserCredential credential;

            // Auth google
            using (var stream = new FileStream(credentialFilePath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { CalendarService.Scope.Calendar },
                    "user",
                    CancellationToken.None).Result;
            }
            // Create the service
            _service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "timesheet_generator"
            });
        }


        public IDictionary<DateTime, EventResourceCalendar> GetEventsPeriodForCalendar(string calendarName, int year, int month)
        {
            var service = _service;

            var calendar = _service.CalendarList.List().Execute().Items.FirstOrDefault(c => c.Summary == calendarName);


            if (calendar == null)
            {
                throw new ArgumentException($"No calendar was found with the name '{calendarName}'");
            }

            // Set the start and end date for the events to retrieve
            var start = new DateTime(year, month, 1, 0, 0, 0, DateTimeKind.Local);
            var end = new DateTime(year, month, DateTime.DaysInMonth(year, month), 23, 59, 59, DateTimeKind.Local);

            var events = new Dictionary<DateTime, EventResourceCalendar>();
            string pageToken = null;

            do
            {
                // Request events for the calendar
                var request = service.Events.List(calendar.Id);
                request.TimeMin = start;
                request.TimeMax = end;
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
                request.PageToken = pageToken;

                var result = request.Execute();
                var items = result.Items;

                if (items != null && items.Count > 0)
                {
                    foreach (var eventItem in items)
                    {
                        if (eventItem.Summary == "WORK")
                        {
                            var startDateTime = DateTime.Parse(eventItem.Start.Date);
                            events[startDateTime] = new EventResourceCalendar(eventItem.Summary, EventTypeCalendar.Work);  
                        }
                        else if (eventItem.Summary == "WORK REMOTE")
                        {
                            var startDateTime = DateTime.Parse(eventItem.Start.Date);
                            events[startDateTime] = new EventResourceCalendar(eventItem.Summary, EventTypeCalendar.WorkRemote);
                        }
                    }
                }

                pageToken = result.NextPageToken;
            } while (pageToken != null);

            return events;
        }
    }
}
