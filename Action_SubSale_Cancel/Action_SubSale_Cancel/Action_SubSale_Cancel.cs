using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_SubSale_Cancel
{
    public class Action_SubSale_Cancel : IPlugin
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
                Entity upSubSale = new Entity(target.LogicalName, target.Id);
                upSubSale["statecode"] = new OptionSetValue(1);    //inactive
                upSubSale["statuscode"] = new OptionSetValue(100000003);  //Cancel
                upSubSale["bsd_canceldate"] = DateTime.UtcNow;
                upSubSale["bsd_canceler"] = new EntityReference("systemuser", context.UserId);
                upSubSale["bsd_cancelreason"] = reason;
                service.Update(upSubSale);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}