using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_CoownerAppendix_Create
{
    public class Plugin_CoownerAppendix_Create : IPlugin
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
                Entity enAppendix = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_type", "bsd_ra", "bsd_spa" }));
                if (!enAppendix.Contains("bsd_type"))
                    return;
                int bsd_type = ((OptionSetValue)enAppendix["bsd_type"]).Value;

                EntityReference refContract = null;
                string logicalName = null;
                if (bsd_type == 100000000 && enAppendix.Contains("bsd_ra"))   //Reservation Contract
                {
                    logicalName = "bsd_reservationcontract";
                    refContract = (EntityReference)enAppendix["bsd_ra"];
                }
                else if (bsd_type == 100000001 && enAppendix.Contains("bsd_spa"))   //Option Entry
                {
                    logicalName = "bsd_optionentry";
                    refContract = (EntityReference)enAppendix["bsd_spa"];
                }

                if (refContract == null)
                    return;

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_coowner"">
                    <filter>
                      <condition attribute=""{logicalName}"" operator=""eq"" value=""{refContract.Id}"" />
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                    </filter>
                    <order attribute=""createdon"" />
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    foreach (var item in rs.Entities)
                    {
                        Entity newCoOwner = new Entity("bsd_coowner");
                        newCoOwner = item;
                        newCoOwner.Attributes.Remove("bsd_coownerid");
                        newCoOwner.Attributes.Remove("ownerid");
                        newCoOwner.Attributes.Remove(logicalName);
                        newCoOwner["bsd_coownerappendix"] = enAppendix.ToEntityReference();
                        newCoOwner.Id = Guid.NewGuid();
                        service.Create(newCoOwner);
                    }
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}