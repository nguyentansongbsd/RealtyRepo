using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_OptionEntry_ChangeHandoverCondition
{
    public class Plugin_OptionEntry_ChangeHandoverCondition : IPlugin
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
                Entity enOE = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_quoteid", "bsd_reservationcontract", "bsd_handovercondition",
                "bsd_detailamount", "bsd_totalamountlessfreight", "bsd_totaltax", "bsd_freightamount"}));

                if (enOE.Contains("bsd_quoteid") || enOE.Contains("bsd_reservationcontract"))  // từ convert
                    return;

                decimal bsd_packagesellingamount = 0;
                decimal bsd_totalamountlessfreight = enOE.Contains("bsd_totalamountlessfreight") ? ((Money)enOE["bsd_totalamountlessfreight"]).Value : 0;
                decimal bsd_totaltax = enOE.Contains("bsd_totaltax") ? ((Money)enOE["bsd_totaltax"]).Value : 0;
                decimal bsd_freightamount = enOE.Contains("bsd_freightamount") ? ((Money)enOE["bsd_freightamount"]).Value : 0;

                Entity upOE = new Entity(enOE.LogicalName, enOE.Id);
                if (enOE.Contains("bsd_handovercondition"))
                {
                    EntityReference refHandover = (EntityReference)enOE["bsd_handovercondition"];
                    Entity enHandover = service.Retrieve(refHandover.LogicalName, refHandover.Id, new ColumnSet(new string[] { "bsd_method", "bsd_amount", "bsd_percent" }));
                    int bsd_method = enHandover.Contains("bsd_method") ? ((OptionSetValue)enHandover["bsd_method"]).Value : -99;
                    if (bsd_method == 100000001)    //Amount
                    {
                        bsd_packagesellingamount = enHandover.Contains("bsd_amount") ? ((Money)enHandover["bsd_amount"]).Value : 0;
                    }
                    else if (bsd_method == 100000002)   //Percent (%)
                    {
                        decimal bsd_percent = enHandover.Contains("bsd_percent") ? (decimal)enHandover["bsd_percent"] / 100 : 0;
                        decimal bsd_detailamount = enOE.Contains("bsd_detailamount") ? ((Money)enOE["bsd_detailamount"]).Value : 0;

                        bsd_packagesellingamount = bsd_detailamount * bsd_percent;
                    }
                }
                else
                {
                    bsd_packagesellingamount = 0;
                }

                upOE["bsd_packagesellingamount"] = new Money(bsd_packagesellingamount);
                upOE["bsd_totalamount"] = new Money(bsd_totalamountlessfreight + bsd_totaltax + bsd_freightamount + bsd_packagesellingamount);
                service.Update(upOE);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}