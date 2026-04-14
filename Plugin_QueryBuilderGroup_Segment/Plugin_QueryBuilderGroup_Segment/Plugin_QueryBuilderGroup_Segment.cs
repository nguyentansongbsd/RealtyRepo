using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_QueryBuilderGroup_Segment
{
    public class Plugin_QueryBuilderGroup_Segment : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 2) return;

                Entity target = (Entity)context.InputParameters["Target"];
                Entity enQBG = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_regardingobjectid"}));
                if (!enQBG.Contains("bsd_regardingobjectid"))
                    return;

                EntityReference refRO = (EntityReference)enQBG["bsd_regardingobjectid"];
                if (refRO.LogicalName != "bsd_segment")
                    return;

                OrganizationRequest req = new OrganizationRequest("bsd_Action_Segment_Dynamic");
                req["Target"] = refRO;
                service.Execute(req);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}