using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_ConstructionMileston_Approve
{
    public class Action_ConstructionMileston_Approve : IPlugin
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

            try
            {
                string data = (string)context.InputParameters["data"];
                string date = (string)context.InputParameters["date"];

                if (string.IsNullOrWhiteSpace(data) || string.IsNullOrWhiteSpace(date))
                    return;

                DateTime selectedDate = RetrieveLocalTimeFromUTCTime(DateTimeOffset.FromUnixTimeMilliseconds(long.Parse(date)).UtcDateTime, service);
                List<Guid> selectedIds = JsonConvert.DeserializeObject<List<Guid>>(data);
                traceService.Trace($"{selectedDate} || {selectedIds}");

                foreach (Guid id in selectedIds)
                {
                    Entity upMilestone = new Entity("bsd_constructionmilestone", id);
                    upMilestone["bsd_actualdate"] = selectedDate;
                    upMilestone["statuscode"] = new OptionSetValue(100000001);  //Completed
                    service.Update(upMilestone);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new Exception("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };
            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
            return response.LocalTime;
        }

        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions =
            {
            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
            }
                }
            }).Entities[0].ToEntity<Entity>();

            return (int?)currentUserSettings.Attributes["timezonecode"];
        }

    }
}
