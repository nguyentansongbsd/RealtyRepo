using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_OptionEntry_Approved
{
    public class Plugin_OptionEntry_Approved : IPlugin
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
                Entity enOE = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_unitnumber" }));
                int status = enOE.Contains("statuscode") ? ((OptionSetValue)enOE["statuscode"]).Value : -99;
                if (status != 100000010)  //Director Approval
                    return;

                EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];
                Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
                upUnit["statuscode"] = new OptionSetValue(100000002);   //Sold
                service.Update(upUnit);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}