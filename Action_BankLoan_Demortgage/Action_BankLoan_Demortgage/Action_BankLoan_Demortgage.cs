using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_BankLoan_Demortgage
{
    public class Action_BankLoan_Demortgage : IPlugin
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

                // up bank loan
                Entity upBankLoan = new Entity(target.LogicalName, target.Id);
                upBankLoan["statuscode"] = new OptionSetValue(100000003);  //Demortgage
                upBankLoan["bsd_demortgageapprovaldate"] = DateTime.UtcNow;
                upBankLoan["bsd_demortgageapprover"] = new EntityReference("systemuser", context.UserId);
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