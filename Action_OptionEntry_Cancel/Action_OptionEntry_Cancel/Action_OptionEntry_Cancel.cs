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

                EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];
                if (enOE.Contains("bsd_reservationcontract"))
                    UpStatus(enOE, refUnit, "bsd_reservationcontract", 100000002, 100000006);
                else if (enOE.Contains("bsd_quoteid"))
                    UpStatus(enOE, refUnit, "bsd_quoteid", 667980002, 100000003);
                else
                    UpStatus(enOE, refUnit);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void UpStatus(Entity enOE, EntityReference refUnit, string fieldContract = null, int? statusContract = null, int? statusUnit = null)
        {
            traceService.Trace($"UpStatus {fieldContract}");

            //up reservation, reservation contract
            if (!string.IsNullOrWhiteSpace(fieldContract))
            {
                EntityReference refContract = (EntityReference)enOE[fieldContract];
                Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
                upContract["statuscode"] = new OptionSetValue((int)statusContract);  //Director Approval
                service.Update(upContract);

                // up unit
                Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
                upUnit["statuscode"] = new OptionSetValue((int)statusUnit);
                service.Update(upUnit);
            }

            string reason = (string)context.InputParameters["reason"];

            // up oe
            Entity upOE = new Entity(enOE.LogicalName, enOE.Id);
            upOE["statuscode"] = new OptionSetValue(100000011);  //Cancel
            upOE["bsd_canceldate"] = DateTime.UtcNow;
            upOE["bsd_canceler"] = new EntityReference("systemuser", context.UserId);
            upOE["bsd_cancelreason"] = reason;
            service.Update(upOE);
        }
    }
}