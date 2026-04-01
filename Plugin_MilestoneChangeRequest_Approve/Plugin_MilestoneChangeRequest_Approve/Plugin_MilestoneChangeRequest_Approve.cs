using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_MilestoneChangeRequest_Approve
{
    public class Plugin_MilestoneChangeRequest_Approve : IPlugin
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
                Entity enRequest = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_progress" }));
                int status = enRequest.Contains("statuscode") ? ((OptionSetValue)enRequest["statuscode"]).Value : -99;
                if (status != 100000001 || !enRequest.Contains("bsd_progress"))    //Approve
                    return;

                UpdatePlannedDate(enRequest);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private EntityCollection GetConstructionMilestone(EntityReference refCP)
        {
            traceService.Trace("GetConstructionMilestone");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_constructionmilestone"">
                <attribute name=""bsd_constructionmilestoneid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_planneddate"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_planneddate"" operator=""not-null"" />
                  <condition attribute=""bsd_constructionprogress"" operator=""eq"" value=""{refCP.Id}"" />
                </filter>
                <order attribute=""bsd_sequence"" />
              </entity>
            </fetch>";
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        private EntityCollection GetMilestoneChangeRequestDetail(Entity enRequest)
        {
            traceService.Trace("GetMilestoneChangeRequestDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_milestonechangerequestdetail"">
                <attribute name=""bsd_milestonechangerequestdetailid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_newplanneddate"" />
                <attribute name=""bsd_constructionmilestone"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_constructionmilestone"" operator=""not-null"" />
                  <condition attribute=""bsd_newplanneddate"" operator=""not-null"" />
                  <condition attribute=""bsd_milestonechangerequest"" operator=""eq"" value=""{enRequest.Id}"" />
                </filter>
              </entity>
            </fetch>";
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        private void UpdatePlannedDate(Entity enRequest)
        {
            traceService.Trace("UpdatePlannedDate");

            EntityReference refCP = (EntityReference)enRequest["bsd_progress"];
            EntityCollection listMilestoneMaster = GetConstructionMilestone(refCP);
            EntityCollection listMilestoneUpdate = GetMilestoneChangeRequestDetail(enRequest);

            var listUpdateMap = listMilestoneUpdate.Entities
                    .ToDictionary(x => ((EntityReference)x["bsd_constructionmilestone"]).Id, x => RetrieveLocalTimeFromUTCTime((DateTime)x["bsd_newplanneddate"], service));

            //merge
            var merged = listMilestoneMaster.Entities.Select(master =>
            {
                var id = master.Id;
                DateTime date = RetrieveLocalTimeFromUTCTime((DateTime)master["bsd_planneddate"], service);

                if (listUpdateMap.ContainsKey(id))
                {
                    date = listUpdateMap[id];
                }

                return new
                {
                    Milestone = master,
                    PlannedDate = date
                };
            }).ToList();

            // check ngày
            for (int i = 1; i < merged.Count; i++)
            {
                var prev = merged[i - 1];
                var current = merged[i];

                var prevDate = prev.PlannedDate;
                var currentDate = current.PlannedDate;

                traceService.Trace($"{current.Milestone.Id} || {currentDate.Date} || {prevDate.Date}");
                if (currentDate.Date < prevDate.Date)
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "planned_date_milestone"));
            }

            // update mistone master
            for (int i = merged.Count - 1; i >= 0; i--)
            {
                var item = merged[i];
                DateTime originalDate = RetrieveLocalTimeFromUTCTime((DateTime)item.Milestone["bsd_planneddate"], service);
                if (originalDate.Date != item.PlannedDate.Date)
                {
                    Entity update = new Entity(item.Milestone.LogicalName, item.Milestone.Id);
                    update["bsd_planneddate"] = item.PlannedDate;
                    service.Update(update);
                }
            }

            //update mistone request detail
            foreach (Entity item in listMilestoneUpdate.Entities)
            {
                Entity enUp = new Entity(item.LogicalName, item.Id);
                enUp["statuscode"] = new OptionSetValue(100000000);   //Approve
                service.Update(enUp);
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