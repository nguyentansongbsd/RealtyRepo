using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_BankLoan_Mortgage
{
    public class Action_BankLoan_Mortgage : IPlugin
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
                upBankLoan["statuscode"] = new OptionSetValue(100000002);  //Mortgage
                upBankLoan["bsd_mortgageapprovaldate"] = DateTime.UtcNow;
                upBankLoan["bsd_mortgageapprover"] = new EntityReference("systemuser", context.UserId);
                service.Update(upBankLoan);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}