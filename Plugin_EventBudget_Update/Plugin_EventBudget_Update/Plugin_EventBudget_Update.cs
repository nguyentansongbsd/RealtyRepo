using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_EventBudget_Update
{
    public class Plugin_EventBudget_Update : IPlugin
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
                Entity enBudget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_event", "statuscode" }));
                if (!enBudget.Contains("bsd_event") || ((OptionSetValue)enBudget["statuscode"]).Value != 100000001) //Approve
                    return;

                EntityReference refEvent = (EntityReference)enBudget["bsd_event"];

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch aggregate=""true"">
                  <entity name=""bsd_eventbudget"">
                    <attribute name=""bsd_budgetamount"" alias=""bsd_budgetamount"" aggregate=""sum"" />
                    <attribute name=""bsd_actualamount"" alias=""bsd_actualamount"" aggregate=""sum"" />
                    <filter>
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                      <condition attribute=""bsd_event"" operator=""eq"" value=""{refEvent.Id}"" />
                      <condition attribute=""statuscode"" operator=""eq"" value=""100000001"" />
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                decimal bsd_budgetamount = 0;
                decimal bsd_actualamount = 0;
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    Entity item = rs[0];
                    if (target.Contains("statuscode") && item.Contains("bsd_budgetamount") && ((AliasedValue)item["bsd_budgetamount"]).Value != null)
                        bsd_budgetamount = ((Money)((AliasedValue)item["bsd_budgetamount"]).Value).Value;

                    if (item.Contains("bsd_actualamount") && ((AliasedValue)item["bsd_actualamount"]).Value != null)
                        bsd_actualamount = ((Money)((AliasedValue)item["bsd_actualamount"]).Value).Value;
                }

                Entity upEvent = new Entity(refEvent.LogicalName, refEvent.Id);
                if (target.Contains("statuscode"))
                    upEvent["bsd_plannedbudget"] = new Money(bsd_budgetamount);
                upEvent["bsd_actualcost"] = new Money(bsd_actualamount);
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