using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_PSAppendix_Cancel
{
    public class Action_PSAppendix_Cancel : IPlugin
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

                // up appendix
                Entity upAppendix = new Entity(target.LogicalName, target.Id);
                upAppendix["statecode"] = new OptionSetValue(1);    //inactive
                upAppendix["statuscode"] = new OptionSetValue(100000002);  //Cancel
                upAppendix["bsd_cancelleddate"] = DateTime.UtcNow;
                upAppendix["bsd_cancelledby"] = new EntityReference("systemuser", context.UserId);
                upAppendix["bsd_cancelledreason"] = reason;
                service.Update(upAppendix);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}