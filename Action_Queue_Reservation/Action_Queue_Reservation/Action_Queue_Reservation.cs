using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IdentityModel.Metadata;
using System.Linq.Expressions;
using System.Runtime.Remoting.Services;
using System.Security.Policy;
using System.Text;
using System.Web.UI.WebControls;
namespace Action_Queue_Reservation
{
    public class Action_Queue_Reservation : IPlugin
    {

        public IOrganizationService service;
        private IOrganizationServiceFactory factory;
        private StringBuilder strbuil = new StringBuilder();
        ITracingService tracingService = null;
        string unitName = "";
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            EntityReference target = (EntityReference)context.InputParameters["Target"];
            string str1 = context.InputParameters["Command"].ToString();
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            Entity queue = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(
                "bsd_phaselaunch", "bsd_pricelist", "bsd_unit", "bsd_project", "bsd_queuingfee", "bsd_customerid", "bsd_queuingfeepaid"));
            Entity updateCurrentQueue = new Entity(target.LogicalName, target.Id);
            updateCurrentQueue["statuscode"] = new OptionSetValue(100000000);
            service.Update(updateCurrentQueue);

            EntityReference unitRef = queue.GetAttributeValue<EntityReference>("bsd_unit");
            Entity unitInfo = service.Retrieve(unitRef.LogicalName, unitRef.Id, new ColumnSet("bsd_name", "bsd_netsaleablearea", "bsd_taxcode", "bsd_maintenancefeespercent", "bsd_maintenancefees"));
            unitName = unitInfo.GetAttributeValue<string>("bsd_name");
            if (unitRef != null)
            {
                
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
                tracingService.Trace("1");
            }
            Entity en_quote = new Entity("bsd_quote");
            en_quote["bsd_opportunityid"] = target; 
            en_quote["bsd_name"] = unitName;
            en_quote["bsd_netusablearea"] = unitInfo.Contains("bsd_netsaleablearea") ? unitInfo["bsd_netsaleablearea"] : Decimal.Zero;
            if (unitInfo.Contains("bsd_taxcode"))
            {
                en_quote["bsd_taxcode"] = unitInfo["bsd_taxcode"];

            }
            var fetchXml1 = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{unitRef.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <attribute name=""bsd_name"" alias=""name"" />
                              <attribute name=""bsd_phaseslaunchid"" alias=""phaseid"" />
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
            EntityCollection rs1 = service.RetrieveMultiple(new FetchExpression(fetchXml1));
            if (rs1.Entities.Count == 1)
            {
                tracingService.Trace("vào if phase_" + rs1.Entities.Count);

                var aliased = (AliasedValue)rs1.Entities[0]["phaseid"];
                Guid phaseId = (Guid)aliased.Value;

                en_quote["bsd_phaseslaunchid"] = new EntityReference("bsd_phaseslaunch", phaseId);
            }
            var fetchXml_pricelist = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <attribute name=""bsd_price"" alias=""prilist_price"" />
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{unitRef.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                          <attribute name=""bsd_name"" alias=""price_name"" />
                          <attribute name=""bsd_pricelevelid"" alias=""price_id"" />
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
            EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml_pricelist));
            if (rs_price.Entities.Count == 1)
            {
                tracingService.Trace("vào if price_" + rs_price.Entities.Count);

                var aliased_price = (AliasedValue)rs_price.Entities[0]["price_id"];
                Guid price_id = (Guid)aliased_price.Value;
                en_quote["bsd_pricelevel"] = new EntityReference("bsd_pricelevel", price_id);
                if (rs_price.Entities[0].Contains("prilist_price"))
                {
                    tracingService.Trace("Có prilist_price");
                    var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
                    Money moneyValue = (Money)aliased_money.Value;

                    en_quote["bsd_detailamount"] = moneyValue;
                    if (unitInfo.Contains("bsd_taxcode"))
                    {
                        tracingService.Trace("Có bsd_taxcode");
                        Entity entity_taxcode = service.Retrieve(((EntityReference)unitInfo["bsd_taxcode"]).LogicalName, ((EntityReference)unitInfo["bsd_taxcode"]).Id, new ColumnSet(true));
                        decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                        decimal taxRate = taxCodeValue / 100.0m;
                        decimal detailAmount = moneyValue.Value;
                        decimal vatAmount = detailAmount * taxRate;
                        en_quote["bsd_vat"] = new Money(vatAmount);
                    }
                }
            }
            if (unitInfo.Contains("bsd_maintenancefeespercent"))
            {
                en_quote["bsd_maintenancefeespercent"] = unitInfo["bsd_maintenancefeespercent"];

            }
            if (unitInfo.Contains("bsd_maintenancefees"))
            {
                en_quote["bsd_maintenancefees"] = unitInfo["bsd_maintenancefees"];

            }
            tracingService.Trace("2");
            en_quote["bsd_reservationtime"] = DateTime.Today;
            if (queue.Contains("bsd_queuingfee")) en_quote["bsd_bookingfee"] = queue.GetAttributeValue<Money>("bsd_queuingfee");
            //if (queue.Contains("bsd_phaselaunch")) en_quote["bsd_phaseslaunchid"] = queue.GetAttributeValue<EntityReference>("bsd_phaselaunch");
            if (queue.Contains("bsd_pricelist")) en_quote["bsd_pricelevel"] = queue.GetAttributeValue<EntityReference>("bsd_pricelist");
            if (queue.Contains("bsd_unit")) en_quote["bsd_unitno"] = queue.GetAttributeValue<EntityReference>("bsd_unit");
            if (queue.Contains("bsd_project")) en_quote["bsd_projectid"] = queue.GetAttributeValue<EntityReference>("bsd_project");
            if (queue.Contains("bsd_customerid")) en_quote["bsd_customerid"] = queue.GetAttributeValue<EntityReference>("bsd_customerid");
            if (queue.Contains("bsd_queuingfeepaid")) en_quote["bsd_totalamountpaid"] = queue.GetAttributeValue<Money>("bsd_queuingfeepaid");
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

