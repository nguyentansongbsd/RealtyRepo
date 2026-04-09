using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_ConstructionMilestone_CreateUpdate
{
    public class Plugin_ConstructionMilestone_CreateUpdate : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        ITracingService traceService = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 2) return;

                Entity target = (Entity)context.InputParameters["Target"];
                Entity enCM = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_milestone", "bsd_constructionprogress",
                    "bsd_completionpercent", "statuscode", "bsd_actualdate", "bsd_sequence", "bsd_planneddate" }));
                if (!enCM.Contains("bsd_constructionprogress"))
                    return;
                EntityReference refCP = (EntityReference)enCM["bsd_constructionprogress"];

                int status = enCM.Contains("statuscode") ? ((OptionSetValue)enCM["statuscode"]).Value : -99;
                traceService.Trace($"status {status}");

                int bsd_sequence = enCM.Contains("bsd_sequence") ? (int)enCM["bsd_sequence"] : 0;
                if (target.Contains("statuscode") && (status == 100000000 || status == 100000001))
                {
                    CheckActualDate(enCM);

                    if (status == 100000001)
                        CheckPrevMileston(enCM, refCP, bsd_sequence);
                }

                if (target.Contains("bsd_milestone"))
                    CheckDuplicateMileston(enCM, refCP);

                if (target.Contains("bsd_completionpercent") && status != 100000001)
                    UpdateMilestonCompleted(enCM);

                if (target.Contains("bsd_planneddate"))
                    CheckPlannedDate(enCM, refCP, bsd_sequence);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void CheckDuplicateMileston(Entity enCM, EntityReference refCP)
        {
            traceService.Trace("CheckDuplicateMileston");

            if (!enCM.Contains("bsd_milestone"))
                return;
            EntityReference refMilestone = (EntityReference)enCM["bsd_milestone"];

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_constructionmilestone"">
                <attribute name=""bsd_constructionmilestoneid"" />
                <attribute name=""bsd_name"" />
                <filter>
                  <condition attribute=""bsd_milestone"" operator=""eq"" value=""{refMilestone.Id}"" />
                  <condition attribute=""bsd_constructionprogress"" operator=""eq"" value=""{refCP.Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 1)
            {
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "exist_milestone"));
            }
        }

        private void UpdateMilestonCompleted(Entity enCM)
        {
            traceService.Trace("UpdateMilestonCompleted");
            decimal bsd_completionpercent = enCM.Contains("bsd_completionpercent") ? (decimal)enCM["bsd_completionpercent"] : 0;
            if (bsd_completionpercent != 100)
                return;

            Entity enUp = new Entity(enCM.LogicalName, enCM.Id);
            enUp["statuscode"] = new OptionSetValue(100000001); //Completed
            service.Update(enUp);
        }

        private void CheckActualDate(Entity enCM)
        {
            traceService.Trace("CheckActualDate");

            if (!enCM.Contains("bsd_actualdate"))
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "actual_date_required_milestone"));
        }

        private void CheckPrevMileston(Entity enCM, EntityReference refCP, int bsd_sequence)
        {
            traceService.Trace("CheckPrevMileston");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_constructionmilestone"">
                <attribute name=""bsd_constructionmilestoneid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_sequence"" />
                <attribute name=""statuscode"" />
                <filter>
                  <condition attribute=""bsd_constructionprogress"" operator=""eq"" value=""{refCP.Id}"" />
                  <condition attribute=""bsd_sequence"" operator=""lt"" value=""{bsd_sequence}"" />
                  <condition attribute=""statuscode"" operator=""not-in"">
                    <value>100000001</value>
                    <value>2</value>
                  </condition>
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "previous_milestone_not_completed"));
            }
        }

        private void CheckPlannedDate(Entity enCM, EntityReference refCP, int bsd_sequence)
        {
            traceService.Trace("CheckPlannedDate");

            if (!enCM.Contains("bsd_planneddate"))
                return;
            DateTime bsd_planneddate = RetrieveLocalTimeFromUTCTime((DateTime)enCM["bsd_planneddate"], service);

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_constructionmilestone"">
                <attribute name=""bsd_constructionmilestoneid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_planneddate"" />
                <filter>
                  <condition attribute=""bsd_constructionprogress"" operator=""eq"" value=""{refCP.Id}"" />
                </filter>
                <filter type=""or"">
                  <filter>
                    <condition attribute=""bsd_sequence"" operator=""lt"" value=""{bsd_sequence}"" />
                    <condition attribute=""bsd_planneddate"" operator=""gt"" value=""{bsd_planneddate.ToShortDateString()}"" />
                  </filter>
                  <filter>
                    <condition attribute=""bsd_sequence"" operator=""gt"" value=""{bsd_sequence}"" />
                    <condition attribute=""bsd_planneddate"" operator=""lt"" value=""{bsd_planneddate.ToShortDateString()}"" />
                  </filter>
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "planned_date_milestone"));
            }
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