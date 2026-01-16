using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq.Expressions;
using System.Text;

namespace Action_ReservationContract_Update
{
    public class Action_ReservationContract_Update : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("=== START Action_ReservationContract_Update ===");
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                tracingService.Trace($"Context Depth: {context.Depth}");
                string str1 = context.InputParameters["Command"].ToString();
                if (context.Depth > 2)
                {
                    tracingService.Trace("Depth > 3 => Stop plugin");
                    return;
                }

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                tracingService.Trace("Getting Target...");
                EntityReference target = (EntityReference)context.InputParameters["Target"];
                tracingService.Trace($"Target Entity: {target.LogicalName}, ID: {target.Id}");

                // Retrieve Quote
                tracingService.Trace("Retrieving Quote record...");
                Entity RA_Contract = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                tracingService.Trace("Quote retrieved successfully.");

                Entity up_RA_Contract = new Entity(RA_Contract.LogicalName, RA_Contract.Id);
                Entity entity_unit = service.Retrieve(((EntityReference)RA_Contract["bsd_unitno"]).LogicalName, ((EntityReference)RA_Contract["bsd_unitno"]).Id, new ColumnSet(true));
                Entity up_unit = new Entity(entity_unit.LogicalName, entity_unit.Id);
                Entity quote = service.Retrieve(((EntityReference)RA_Contract["bsd_quoteid"]).LogicalName, ((EntityReference)RA_Contract["bsd_quoteid"]).Id, new ColumnSet(true));
                Entity up_quote = new Entity(entity_unit.LogicalName, entity_unit.Id);
                if (str1 == "cancel")
                {
                    checkpayment(RA_Contract.Id, service);
                    if (!RA_Contract.Contains("bsd_quoteid"))
                    {
                        tracingService.Trace("vào if cancel_RAContract");
                        up_RA_Contract["statuscode"] = new OptionSetValue(100000003);//Canceled
                        //up_RA_Contract["bsd_canceldate"] = DateTime.Today;
                        //up_RA_Contract["bsd_canceller"] = new EntityReference("systemuser", context.UserId);
                        service.Update(up_RA_Contract);


                        up_unit["statuscode"] = new OptionSetValue(100000000);
                        service.Update(up_unit);
                    }
                    else
                    {
                        tracingService.Trace("vào else cancel case update khi lên từ cọc");
                        up_unit["statuscode"] = new OptionSetValue(100000003);//Deposited
                        service.Update(up_unit);
                        up_quote["statuscode"] = new OptionSetValue(667980002);//Director Approval
                        //up_RA_Contract["bsd_canceldate"] = DateTime.Today;
                        //up_RA_Contract["bsd_canceller"] = new EntityReference("systemuser", context.UserId);
                        service.Update(up_quote);
                        up_RA_Contract["statuscode"] = new OptionSetValue(100000003);//Canceled
                        service.Update(up_RA_Contract);
                    }

                }



            }
            catch (Exception ex)
            {
                // Trace toàn bộ lỗi
                tracingService.Trace("=== ERROR Action_UpdateQuote ===");
                tracingService.Trace("Message: " + ex.Message);
                tracingService.Trace("StackTrace: " + ex.StackTrace);

                // Ném lỗi ra ngoài để Action bắt được
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void checkpayment(Guid RA_Contract, IOrganizationService service)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_payment"">
                <filter>
                  <condition attribute=""statuscode"" operator=""in"">
                    <value>{1}</value>
                    <value>{100000000}</value>
                  </condition>
                  <condition attribute=""bsd_quotationreservation"" operator=""eq"" value=""{RA_Contract}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (rs.Entities.Count > 0)
            {
                throw new InvalidPluginExecutionException("The transaction has an invalid receipt. Please check again.");
            }
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
