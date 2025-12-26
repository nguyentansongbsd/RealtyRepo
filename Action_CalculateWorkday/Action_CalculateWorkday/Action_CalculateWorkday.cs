using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_CalculateWorkday
{
    public class Action_CalculateWorkday : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("start");

            string startDateInput = context.InputParameters["startDate"].ToString();
            traceService.Trace("startDateInput " + startDateInput);
            DateTime startDate = DateTime.ParseExact(startDateInput, "d/M/yyyy", CultureInfo.InvariantCulture);
            int workDaysToAdd = (int)context.InputParameters["workDaysToAdd"];
            int workingDayType = (int)context.InputParameters["workingDayType"];
            traceService.Trace("workDaysToAdd " + workDaysToAdd);
            traceService.Trace("workingDayType " + workingDayType);
            List<DateTime> holidays = new List<DateTime>();
            if (context.InputParameters.Contains("holidays") && !string.IsNullOrWhiteSpace((string)context.InputParameters["holidays"]))
            {
                traceService.Trace("...");
                holidays = context.InputParameters["holidays"].ToString()
                            .Split(',')
                            .Select(date => Convert.ToDateTime(date))
                            .ToList();
            }
            else
                holidays = GetHolidays();
            traceService.Trace("holidays " + holidays.Count);

            foreach (var date in holidays)
            {
                traceService.Trace(date.ToString());
            }

            DateTime? resultDate = CalculateWorkday(startDate.Date, workDaysToAdd, holidays, workingDayType);
            traceService.Trace("resultDate " + resultDate);
            context.OutputParameters["resultDate"] = resultDate?.ToShortDateString();
        }

        private List<DateTime> GetHolidays()
        {
            traceService.Trace("GetHolidays");
            List<DateTime> holidays = new List<DateTime>();
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""calendar"">
                <filter>
                  <condition attribute=""type"" operator=""eq"" value=""2"" />
                </filter>
                <link-entity name=""calendarrule"" from=""calendarid"" to=""calendarid"">
                  <attribute name=""name"" alias=""namerule"" />
                  <attribute name=""effectiveintervalstart"" alias=""startrule"" />
                  <attribute name=""effectiveintervalend"" alias=""endrule"" />
                  <attribute name=""duration"" alias=""durationrule"" />
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    int duration = item.Contains("durationrule") ? ((int)((AliasedValue)item["durationrule"]).Value / 1440) : 0;
                    DateTime dateHolidayStart = RetrieveLocalTimeFromUTCTime((DateTime)((AliasedValue)item["startrule"]).Value, service);
                    for (int i = 0; i < duration; i++)
                    {
                        holidays.Add(dateHolidayStart.AddDays(i).Date);
                    }
                }
            }

            holidays = holidays.Distinct().ToList();
            return holidays;
        }

        private bool IsWorkingDay(int workingDayType, DayOfWeek day)
        {
            switch (workingDayType)
            {
                case 100000000: // Ngày làm việc từ thứ 2 đến thứ 6
                    return DayOfWeek.Monday <= day && day <= DayOfWeek.Friday;
                case 100000001: // Ngày làm việc từ thứ 2 đến thứ 7
                    return DayOfWeek.Monday <= day && day <= DayOfWeek.Saturday;
                case 100000002: // Ngày làm việc từ thứ 2 đến Chủ Nhật
                    return true;
            }
            return false;
        }

        public DateTime? CalculateWorkday(DateTime startDate, int workDaysToAdd, List<DateTime> holidays, int workingDayType)
        {
            traceService.Trace("startDate " + startDate);
            // Điều chỉnh ngày bắt đầu cho trường hợp cuối tuần
            startDate = AdjustStartDate(startDate, workingDayType);
            int workDaysInWeek = GetWorkDaysInWeek(workingDayType);
            if (workDaysInWeek == 0) return null;

            // Tính số tuần đầy đủ và số ngày làm việc thừa
            int fullWeeks = workDaysToAdd / workDaysInWeek;
            int extraWorkDays = workDaysToAdd % workDaysInWeek;
            int estimatedDays = fullWeeks * 7;  // Bắt đầu từ số tuần đầy đủ

            // Tìm thêm số ngày dư
            DateTime estimatedDate = startDate.AddDays(estimatedDays);
            traceService.Trace("startDate new " + startDate);
            traceService.Trace("estimatedDays " + estimatedDays);
            traceService.Trace("estimatedDate " + estimatedDate);
            traceService.Trace("extraWorkDays " + extraWorkDays);

            List<DateTime> holidaysWithinRange = holidays
            .Where(h => h > startDate && h <= estimatedDate && IsWorkingDay(workingDayType, h.DayOfWeek))
            .ToList();
            traceService.Trace("holidays found count " + holidaysWithinRange.Count);
            extraWorkDays += holidaysWithinRange.Count;
            traceService.Trace("extraWorkDays new " + extraWorkDays);

            while (extraWorkDays > 0)
            {
                estimatedDate = estimatedDate.AddDays(1);
                if (IsWorkingDay(workingDayType, estimatedDate.DayOfWeek) && !holidays.Contains(estimatedDate))
                {
                    extraWorkDays--;
                }
            }

            return estimatedDate;
        }

        private DateTime AdjustStartDate(DateTime startDate, int workingDayType)
        {
            switch (workingDayType)
            {
                case 100000000: // Ngày làm việc từ thứ 2 đến thứ 6
                    if (startDate.DayOfWeek == DayOfWeek.Saturday)
                    {
                        return startDate.AddDays(-1); // Thứ 7 chuyển thành thứ 6
                    }
                    else if (startDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        return startDate.AddDays(-2); // CN chuyển thành thứ 6
                    }
                    break;
                case 100000001: // Ngày làm việc từ thứ 2 đến thứ 7
                    if (startDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        return startDate.AddDays(-1); // CN chuyển thành thứ 7
                    }
                    break;
            }
            return startDate;
        }

        private int GetWorkDaysInWeek(int workingDayType)
        {
            // số ngày làm việc trong tuần
            switch (workingDayType)
            {
                case 100000000: return 5; // Thứ Hai đến thứ Sáu
                case 100000001: return 6; // Thứ Hai đến thứ Bảy
                case 100000002: return 7; // Thứ Hai đến Chủ Nhật
            }
            return 0;
        }

        public DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService ser)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(ser);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            LocalTimeFromUtcTimeResponse response = (LocalTimeFromUtcTimeResponse)ser.Execute(request);
            return response.LocalTime;
            //var utcTime = utcTime.ToString("MM/dd/yyyy HH:mm:ss");
            //var localDateOnly = response.LocalTime.ToString("dd-MM-yyyy");
        }
        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            }).Entities[0].ToEntity<Entity>();

            return (int?)currentUserSettings.Attributes["timezonecode"];
        }
    }
}
