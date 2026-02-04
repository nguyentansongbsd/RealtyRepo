using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_OptionEntry_Cancel
{
    public class Action_OptionEntry_Cancel : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        IPluginExecutionContext context = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 1) return;

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity enOE = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_unitnumber", "bsd_quoteid", "bsd_reservationcontract" }));

                if (!enOE.Contains("bsd_unitnumber"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_unitnumber"));

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_payment"">
                    <attribute name=""bsd_paymentid"" />
                    <attribute name=""bsd_name"" />
                    <filter>
                      <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{target.Id}"" />
                      <condition attribute=""statuscode"" operator=""not-in"">
                        <value>100000000</value>
                      </condition>
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "spa_existing_receipts"));
                }

                EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];
                if (enOE.Contains("bsd_reservationcontract"))   //hđcs
                    UpStatus(enOE, refUnit, "bsd_reservationcontract", 100000000, 100000006);
                else if (enOE.Contains("bsd_quoteid"))  //đặt cọc
                    UpStatus(enOE, refUnit, "bsd_quoteid", 667980008, 100000003);
                else  //sản phẩm
                    UpStatus(enOE, refUnit, null, 0, 100000000);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void UpStatus(Entity enOE, EntityReference refUnit, string fieldContract, int statusContract, int statusUnit)
        {
            traceService.Trace($"UpStatus {fieldContract}");

            //up reservation, reservation contract
            if (!string.IsNullOrWhiteSpace(fieldContract))
            {
                EntityReference refContract = (EntityReference)enOE[fieldContract];
                Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
                upContract["statecode"] = new OptionSetValue(0);    //active
                upContract["statuscode"] = new OptionSetValue(statusContract);  //Director Approval
                service.Update(upContract);
            }

            string reason = (string)context.InputParameters["reason"];

            // up oe
            Entity upOE = new Entity(enOE.LogicalName, enOE.Id);
            upOE["statecode"] = new OptionSetValue(1);    //inactive
            upOE["statuscode"] = new OptionSetValue(100000012);  //Cancel
            upOE["bsd_canceldate"] = DateTime.UtcNow;
            upOE["bsd_canceler"] = new EntityReference("systemuser", context.UserId);
            upOE["bsd_cancelreason"] = reason;
            service.Update(upOE);

            // up unit
            Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
            upUnit["statuscode"] = new OptionSetValue(statusUnit);
            service.Update(upUnit);
        }
    }
}