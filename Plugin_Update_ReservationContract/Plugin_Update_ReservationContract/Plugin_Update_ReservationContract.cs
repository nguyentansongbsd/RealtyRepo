using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_Update_ReservationContract
{
    public class Plugin_Update_ReservationContract : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity target = context.InputParameters["Target"] as Entity;

            Entity Re_contract = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
            Entity up_Re_contract = new Entity(Re_contract.LogicalName, Re_contract.Id);
            if (context.Depth > 3)
            {
                return;
            }
            ////
            //if (Re_contract.Contains("bsd_unitno") && Re_contract.Contains("bsd_pricelevel"))
            //{
            //    Entity enUnit = service.Retrieve(((EntityReference)Re_contract["bsd_unitno"]).LogicalName, ((EntityReference)quote["bsd_unitno"]).Id, new ColumnSet(true));
            //    Entity en_price = service.Retrieve(((EntityReference)Re_contract["bsd_pricelevel"]).LogicalName, ((EntityReference)quote["bsd_pricelevel"]).Id, new ColumnSet(true));

            //    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            //<fetch distinct=""true"">
            //  <entity name=""bsd_productpricelevel"">
            //    <attribute name=""bsd_price"" alias=""prilist_price"" />
            //    <filter>
            //      <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
            //    </filter>
            //    <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
            //      <filter>
            //        <condition attribute=""bsd_pricelevelid"" operator=""eq"" value=""{en_price.Id}"" />
            //      </filter>
            //    </link-entity>
            //  </entity>
            //</fetch>";
            //    EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml));
            //    if (rs_price.Entities.Count > 0)
            //    {
            //        trace.Trace("vào if price_" + rs_price.Entities.Count);
            //        if (rs_price.Entities[0].Contains("prilist_price"))
            //        {
            //            trace.Trace("Có prilist_price");
            //            var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
            //            Money moneyValue = (Money)aliased_money.Value;

            //            up_Re_contract["bsd_detailamount"] = moneyValue;
            //            //if (enUnit.Contains("bsd_taxcode"))
            //            //{
            //            //    trace.Trace("Có bsd_taxcode");
            //            //    Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
            //            //    decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
            //            //    decimal taxRate = taxCodeValue / 100.0m;
            //            //    decimal detailAmount1 = moneyValue.Value;
            //            //    decimal vatAmount = detailAmount1 * taxRate;
            //            //    up_quote["bsd_vat"] = new Money(vatAmount);
            //                service.Update(up_Re_contract);
            //            }
            //        }
            //    }
            //}
            ///
            decimal discountAmount = 0;
            decimal detailAmount = 0;
            decimal bsd_landvaluededuction = 0;
            decimal bsd_packagesellingamount = 0;
            detailAmount = Re_contract.Contains("bsd_detailamount") ? ((Money)Re_contract["bsd_detailamount"]).Value : 0;
            discountAmount = Re_contract.Contains("bsd_discountamount") ? ((Money)Re_contract["bsd_discountamount"]).Value : 0;
            bsd_landvaluededuction = Re_contract.Contains("bsd_landvaluededuction") ? ((Money)Re_contract["bsd_landvaluededuction"]).Value : 0;
            bsd_packagesellingamount = Re_contract.Contains("bsd_packagesellingamount") ? ((Money)Re_contract["bsd_packagesellingamount"]).Value : 0;
            if (Re_contract.Contains("bsd_handovercondition"))
            {
                Entity handover = service.Retrieve(((EntityReference)Re_contract["bsd_handovercondition"]).LogicalName, ((EntityReference)Re_contract["bsd_handovercondition"]).Id, new ColumnSet(true));
                int bsd_method = handover.Contains("bsd_method") ? ((OptionSetValue)handover["bsd_method"]).Value : 0;
                if (bsd_method == 100000001)
                {
                    bsd_packagesellingamount = handover.Contains("bsd_amount") ? ((Money)handover["bsd_amount"]).Value : 0;
                    up_Re_contract["bsd_packagesellingamount"] = new Money(bsd_packagesellingamount);
                }
                else if (bsd_method == 100000002)
                {
                    decimal bsd_percent = handover.Contains("bsd_percent") ? (decimal)handover["bsd_percent"] : 0;
                    bsd_packagesellingamount = bsd_percent / 100.0m * (detailAmount - discountAmount);
                    up_Re_contract["bsd_packagesellingamount"] = new Money(bsd_packagesellingamount);
                }
            }   
            
            decimal totalamountlessfreight = detailAmount - discountAmount + bsd_packagesellingamount;
            up_Re_contract["bsd_totalamountlessfreight"] = new Money(totalamountlessfreight);

            Entity entity_taxcode = service.Retrieve(((EntityReference)Re_contract["bsd_taxcode"]).LogicalName, ((EntityReference)Re_contract["bsd_taxcode"]).Id, new ColumnSet("bsd_value"));
            decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
            decimal taxRate = taxCodeValue / 100.0m;
            up_Re_contract["bsd_totaltax"] = new Money((totalamountlessfreight - bsd_landvaluededuction) * taxRate);
            up_Re_contract["bsd_totalamountlessfreightaftervat"] = new Money(totalamountlessfreight + ((totalamountlessfreight - bsd_landvaluededuction) * taxRate));
            Entity entity_unit = service.Retrieve(((EntityReference)Re_contract["bsd_unitno"]).LogicalName, ((EntityReference)Re_contract["bsd_unitno"]).Id, new ColumnSet("bsd_maintenancefeespercent"));
            decimal percen1 = entity_unit.Contains("bsd_maintenancefeespercent") ? (decimal)entity_unit["bsd_maintenancefeespercent"] : 0;
            decimal taxper = percen1 / 100.0m;
            up_Re_contract["bsd_freightamount"] = new Money(taxper * totalamountlessfreight);
            
            up_Re_contract["bsd_totalamount"] = new Money(totalamountlessfreight + (totalamountlessfreight - bsd_landvaluededuction) * taxRate +taxper * totalamountlessfreight);
            service.Update(up_Re_contract);

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

