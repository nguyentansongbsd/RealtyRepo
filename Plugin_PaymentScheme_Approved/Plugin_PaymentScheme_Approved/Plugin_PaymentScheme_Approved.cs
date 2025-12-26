using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PaymentScheme_Approved
{
    public class Plugin_PaymentScheme_Approved : IPlugin
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
                Entity enPS = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_phaselaunch" }));
                int status = enPS.Contains("statuscode") ? ((OptionSetValue)enPS["statuscode"]).Value : -99;
                if (status != 100000000 || !enPS.Contains("bsd_phaselaunch"))  //Approved
                    return;

                EntityReference refPL = (EntityReference)enPS["bsd_phaselaunch"];
                var relativeEntity = new EntityReferenceCollection { new EntityReference(refPL.LogicalName, refPL.Id) };
                Relationship relationship = new Relationship("bsd_bsd_phaseslaunch_bsd_paymentscheme");
                service.Associate(enPS.LogicalName, enPS.Id, relationship, relativeEntity);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}