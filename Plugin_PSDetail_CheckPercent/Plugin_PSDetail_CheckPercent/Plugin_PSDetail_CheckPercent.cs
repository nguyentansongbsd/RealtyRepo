using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PSDetail_CheckPercent
{
    public class Plugin_PSDetail_CheckPercent : IPlugin
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
                Entity enIns = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_paymentscheme", "bsd_pricetype" }));
                if (!enIns.Contains("bsd_paymentscheme"))
                    throw new InvalidPluginExecutionException("There is no 'Payment Scheme'. Please check again.");

                if (!enIns.Contains("bsd_pricetype"))
                    throw new InvalidPluginExecutionException("There is no 'Price Type'. Please check again.");

                EntityReference refPS = (EntityReference)enIns["bsd_paymentscheme"];
                int bsd_pricetype = ((OptionSetValue)enIns["bsd_pricetype"]).Value;

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch aggregate=""true"">
                  <entity name=""bsd_paymentschemedetailmaster"">
                    <attribute name=""bsd_amountpercent"" alias=""bsd_amountpercent"" aggregate=""sum"" />
                    <filter>
                      <condition attribute=""bsd_paymentscheme"" operator=""eq"" value=""{refPS.Id}"" />
                      <condition attribute=""bsd_pricetype"" operator=""eq"" value=""{bsd_pricetype}"" />
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    decimal sumPercent = 0;
                    if (((AliasedValue)rs[0]["bsd_amountpercent"]).Value != null)
                        sumPercent = (decimal)((AliasedValue)rs[0]["bsd_amountpercent"]).Value;

                    if (sumPercent > 100)
                        throw new InvalidPluginExecutionException("The total percentage is greater than 100%. Please check again.");
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