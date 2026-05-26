using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Action_BankLoan_Cancel
{
    public class Action_BankLoan_Cancel : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 1) return;

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                string reason = (string)context.InputParameters["reason"];

                // up bank loan
                Entity upBankLoan = new Entity(target.LogicalName, target.Id);
                upBankLoan["statecode"] = new OptionSetValue(1);    //inactive
                upBankLoan["statuscode"] = new OptionSetValue(100000004);  //Cancel
                upBankLoan["bsd_canceldate"] = DateTime.UtcNow;
                upBankLoan["bsd_cancelledby"] = new EntityReference("systemuser", context.UserId);
                upBankLoan["bsd_cancelreason"] = reason;
                service.Update(upBankLoan);

                // up unit
                Entity enBankLoan = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_units" }));
                EntityReference refUnit = (EntityReference)enBankLoan["bsd_units"];

                Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
                upUnit["bsd_bankloan"] = false;
                service.Update(upUnit);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}