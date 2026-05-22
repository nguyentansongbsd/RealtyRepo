using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Termination_CreateUpdate
{
    public class Plugin_Termination_CreateUpdate : IPlugin
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
                Entity enTermination = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_units" }));

                if (context.MessageName == "Create")
                    UpdateUnit(enTermination);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void UpdateUnit(Entity enTermination)
        {
            traceService.Trace("UpdateUnit");

            EntityReference refUnit = (EntityReference)enTermination["bsd_units"];

            Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
            upUnit["bsd_isterminate"] = true;
            service.Update(upUnit);
        }
    }
}