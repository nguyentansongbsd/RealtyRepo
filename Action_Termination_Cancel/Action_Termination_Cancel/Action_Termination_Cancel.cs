using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Termination_Cancel
{
    public class Action_Termination_Cancel : IPlugin
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
                string reason = (string)context.InputParameters["reason"];

                // up oe
                Entity upTermination = new Entity(target.LogicalName, target.Id);
                upTermination["statecode"] = new OptionSetValue(1);    //inactive
                upTermination["statuscode"] = new OptionSetValue(100000003);  //Cancel
                upTermination["bsd_canceldate"] = DateTime.UtcNow;
                upTermination["bsd_canceler"] = new EntityReference("systemuser", context.UserId);
                upTermination["bsd_cancelreason"] = reason;
                service.Update(upTermination);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}