using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IdentityModel.Metadata;
using System.Linq.Expressions;
using System.Runtime.Remoting.Services;
using System.Security.Policy;
using System.Text;
namespace Action_Queue_Reservation
{
    public class Action_Queue_Reservation : IPlugin
    {

        public IOrganizationService service;
        private IOrganizationServiceFactory factory;
        private StringBuilder strbuil = new StringBuilder();
        ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            EntityReference target = (EntityReference)context.InputParameters["Target"];
            string str1 = context.InputParameters["Command"].ToString();
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            Entity queue = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(
                "bsd_phaselaunch", "bsd_pricelist", "bsd_unit", "bsd_project"));
            Entity updateCurrentQueue = new Entity(target.LogicalName, target.Id);
            updateCurrentQueue["statuscode"] = new OptionSetValue(100000000);
            service.Update(updateCurrentQueue);

            EntityReference unitRef = queue.GetAttributeValue<EntityReference>("bsd_unit");
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""statuscode"" />
                <filter>
                  <condition attribute=""bsd_unit"" operator=""eq"" value=""{unitRef.Id}"" />
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{target.Id}"" />
                  <condition attribute=""statuscode"" operator=""ne"" value=""{100000005}"" />
                  <condition attribute=""statuscode"" operator=""ne"" value=""{100000001}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            foreach (var q in rs.Entities)
            {
                Entity updateOther = new Entity(target.LogicalName, q.Id);
                updateOther["statuscode"] = new OptionSetValue(100000002);
                service.Update(updateOther);
            }
            Entity updateUnit = new Entity(unitRef.LogicalName, unitRef.Id);
            updateUnit["statuscode"] = new OptionSetValue(100000003);
            service.Update(updateUnit);

            Entity en_quote = new Entity("bsd_quote");
            en_quote["statuscode"] = new OptionSetValue(100000000);
            en_quote["bsd_phaseslaunchid"] = queue.GetAttributeValue<EntityReference>("bsd_phaselaunch");
            en_quote["bsd_pricelevel"] = queue.GetAttributeValue<EntityReference>("bsd_pricelist");
            en_quote["bsd_unitno"] = queue.GetAttributeValue<EntityReference>("bsd_unit");
            en_quote["bsd_projectid"] = queue.GetAttributeValue<EntityReference>("bsd_project");
            Guid guid = service.Create(en_quote);
            context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + guid.ToString() + "'}";
        }
        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
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
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            }).Entities[0].ToEntity<Entity>();
            return (int?)currentUserSettings.Attributes["timezonecode"];
        }
    }
}

