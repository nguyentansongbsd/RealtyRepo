using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_SubSale_CoOwner
{
    public class Plugin_SubSale_CoOwner : IPlugin
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
                Entity enSubSale = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_type", "bsd_quote", "bsd_reservationcontract", "bsd_optionentry" }));

                if (!enSubSale.Contains("bsd_type"))
                    return;
                int bsd_type = ((OptionSetValue)enSubSale["bsd_type"]).Value;

                string logicalName = null;
                if (bsd_type == 100000000 && enSubSale.Contains("bsd_quote"))   //Deposit
                {
                    logicalName = "bsd_reservation";
                }
                else if (bsd_type == 100000001 && enSubSale.Contains("bsd_reservationcontract"))   //Reservation Contract
                {
                    logicalName = "bsd_reservationcontract";
                }
                else if (bsd_type == 100000002 && enSubSale.Contains("bsd_optionentry"))   //Option Entry
                {
                    logicalName = "bsd_optionentry";
                }

                if (logicalName == null)
                    return;
                EntityReference refContract = (EntityReference)enSubSale[logicalName];

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
                        Entity newSubSale = new Entity("bsd_assign");
                        newSubSale = item;
                        newSubSale.Attributes.Remove("bsd_coownerid");
                        newSubSale.Attributes.Remove("ownerid");
                        newSubSale.Attributes.Remove(logicalName);
                        newSubSale["bsd_subsale"] = enSubSale.ToEntityReference();
                        newSubSale.Id = Guid.NewGuid();
                        service.Create(newSubSale);
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