using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_EventRegistration_CountSts
{
    public class Plugin_EventRegistration_CountSts : IPlugin
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
                Entity enER = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_event" }));
                if (!enER.Contains("bsd_event"))
                    return;

                EntityReference refEvent = (EntityReference)enER["bsd_event"];

                var query = new QueryExpression("bsd_eventregistration");
                query.ColumnSet.AddColumns("statuscode", "bsd_checkinstatus");
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("bsd_event", ConditionOperator.Equal, refEvent.Id);
                var query_Or = new FilterExpression(LogicalOperator.Or);
                query.Criteria.AddFilter(query_Or);
                if (target.Contains("statuscode"))
                    query_Or.AddCondition("statuscode", ConditionOperator.Equal, 100000001);
                if (target.Contains("bsd_checkinstatus"))
                    query_Or.AddCondition("bsd_checkinstatus", ConditionOperator.Equal, 100000001);
                EntityCollection rs = service.RetrieveMultiple(query);
                int cntStatus = 0;
                int cntCheckin = 0;
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {

                    foreach (var item in rs.Entities)
                    {
                        if (target.Contains("statuscode") && ((OptionSetValue)item["statuscode"]).Value == 100000001)
                            cntStatus++;

                        if (target.Contains("bsd_checkinstatus") && item.Contains("bsd_checkinstatus") && ((OptionSetValue)item["bsd_checkinstatus"]).Value == 100000001)
                            cntCheckin++;
                    }
                }

                Entity upEvent = new Entity(refEvent.LogicalName, refEvent.Id);
                if (target.Contains("statuscode"))
                    upEvent["bsd_registrationcount"] = cntStatus;
                if (target.Contains("bsd_checkinstatus"))
                    upEvent["bsd_checkincount"] = cntCheckin;
                service.Update(upEvent);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}