using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_OptionEntry_ApproveDebt
{
    public class Action_OptionEntry_ApproveDebt : IPlugin
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
                throw new InvalidPluginExecutionException("test");
                EntityReference target = (EntityReference)context.InputParameters["Target"];

                Entity upOE = new Entity(target.LogicalName, target.Id);
                upOE["bsd_debtapprover"] = new EntityReference("systemuser", context.UserId);
                upOE["bsd_debtapprovaldate"] = DateTime.UtcNow;
                service.Update(upOE);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

    }
}