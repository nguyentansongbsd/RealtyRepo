using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_OptionEntry_Sign
{
    public class Action_OptionEntry_Sign : IPlugin
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

                // up oe
                Entity upOE = new Entity(target.LogicalName, target.Id);
                upOE["statuscode"] = new OptionSetValue(100000013);  //Signed
                upOE["bsd_tellersdate"] = DateTime.UtcNow;
                upOE["bsd_signedby"] = new EntityReference("systemuser", context.UserId);
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