using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_FollowUpList_Cancel
{
    public class Action_FollowUpList_Cancel : IPlugin
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

                // up ful
                Entity upFUL = new Entity(target.LogicalName, target.Id);
                upFUL["statecode"] = new OptionSetValue(1);    //inactive
                upFUL["statuscode"] = new OptionSetValue(100000003);  //Cancel
                upFUL["bsd_canceldate"] = DateTime.UtcNow;
                upFUL["bsd_canceler"] = new EntityReference("systemuser", context.UserId);
                upFUL["bsd_cancelreason"] = reason;
                service.Update(upFUL);

                // up unit
                Entity enFUL = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_units" }));
                EntityReference refUnit = (EntityReference)enFUL["bsd_units"];

                Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
                upUnit["bsd_isfollowuplist"] = false;
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