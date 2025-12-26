using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;

namespace Action_UpdateQuote
{
    public class Action_UpdateQuote : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("=== START Action_UpdateQuote ===");

            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                tracingService.Trace($"Context Depth: {context.Depth}");
                string str1 = context.InputParameters["Command"].ToString();
                if (context.Depth > 3)
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
                Entity quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                tracingService.Trace("Quote retrieved successfully.");

                Entity up_quote = new Entity(quote.LogicalName, quote.Id);
                
                Entity entity_unit = service.Retrieve(((EntityReference)quote["bsd_unitno"]).LogicalName, ((EntityReference)quote["bsd_unitno"]).Id, new ColumnSet(true));
                Entity up_unit = new Entity(entity_unit.LogicalName, entity_unit.Id);
                if (str1 == "confirm")
                {
                    // Fetch Tiến độ thanh toán
                    tracingService.Trace("Running FetchXML to get bsd_paymentschemedetail...");
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch>
                      <entity name=""bsd_paymentschemedetail"">
                        <attribute name=""bsd_name"" />
                        <filter>
                          <condition attribute=""bsd_quotation"" operator=""eq"" value=""{target.Id}"" />
                          <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                        </filter>
                      </entity>
                    </fetch>";

                    tracingService.Trace("FetchXML: " + fetchXml);

                    EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));


                    if (rs.Entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException("Vui lòng tạo Tiến độ thanh toán");
                    }

                    // Update fields
                    up_quote["bsd_confirmdate"] = DateTime.Today;
                    up_quote["bsd_confirmer"] = new EntityReference("systemuser", context.UserId);
                    up_quote["statuscode"] = new OptionSetValue(100000000);
                    service.Update(up_quote);
                    tracingService.Trace("Quote updated successfully");
                }
                if (str1 == "cancel")
                {
                    up_quote["statuscode"] = new OptionSetValue(100000002);
                    up_quote["bsd_canceldate"] = DateTime.Today;
                    up_quote["bsd_canceller"] = new EntityReference("systemuser", context.UserId);
                    service.Update(up_quote);
                    OptionSetValue statusCode = entity_unit.GetAttributeValue<OptionSetValue>("statuscode");
                    if (statusCode != null && statusCode.Value != 100000004)
                    {
                        up_unit["statuscode"] = new OptionSetValue(100000000);
                        service.Update(up_unit);
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
