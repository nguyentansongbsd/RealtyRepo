using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_PhasesLaunch_StartStopSelling
{
    public class Action_PhasesLaunch_StartStopSelling : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("start");

            EntityReference refPL = (EntityReference)context.InputParameters["Target"];
            int type = (int)context.InputParameters["type"];
            traceService.Trace($"type: {type}");

            // 1: ngưng bán, 0: mở bán
            Entity upPL = new Entity(refPL.LogicalName, refPL.Id);
            upPL["bsd_stopselling"] = type == 1 ? true : false;
            service.Update(upPL);

            traceService.Trace("done");
        }
    }
}