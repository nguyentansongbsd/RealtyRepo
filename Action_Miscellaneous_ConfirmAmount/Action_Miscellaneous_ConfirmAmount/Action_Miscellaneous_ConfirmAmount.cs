using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Miscellaneous_ConfirmAmount
{
    public class Action_Miscellaneous_ConfirmAmount : IPlugin
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

                Entity upTarget = new Entity(target.LogicalName, target.Id);
                upTarget["statuscode"] = new OptionSetValue(100000000);  //Not Paid
                service.Update(upTarget);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}