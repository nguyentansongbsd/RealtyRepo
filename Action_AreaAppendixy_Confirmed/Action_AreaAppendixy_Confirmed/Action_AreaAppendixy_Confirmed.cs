using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Security.Policy;

namespace Action_AreaAppendixy_Confirmed
{
    public class Action_AreaAppendixy_Confirmed : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService traceS = null;
        EntityReference target = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                target = (EntityReference)context.InputParameters["Target"];
                traceS = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceS.Trace($"start {target.Id}");
                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                Entity enUp = new Entity(target.LogicalName, target.Id);
                enUp["statuscode"] = new OptionSetValue(100000001);
                enUp["bsd_confirmedby"] = new EntityReference("systemuser", context.UserId);
                enUp["bsd_confirmeddate"] = DateTime.Now;
                service.Update(enUp);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}