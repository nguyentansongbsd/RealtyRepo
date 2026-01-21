using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Discount_Approved
{
    public class Plugin_Discount_Approved : IPlugin
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
                Entity enDiscount = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_phaselaunch", "bsd_priority" }));
                int status = enDiscount.Contains("statuscode") ? ((OptionSetValue)enDiscount["statuscode"]).Value : -99;
                if (status != 100000000 || !enDiscount.Contains("bsd_phaselaunch"))  //Approved
                    return;

                Entity newDR = new Entity("bsd_discountrule");
                newDR["bsd_discount"] = enDiscount.ToEntityReference();
                newDR["bsd_priority"] = enDiscount.Contains("bsd_priority") ? enDiscount["bsd_priority"] : null;
                newDR["bsd_phaseslaunch"] = enDiscount["bsd_phaselaunch"];
                newDR.Id = Guid.NewGuid();
                service.Create(newDR);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}