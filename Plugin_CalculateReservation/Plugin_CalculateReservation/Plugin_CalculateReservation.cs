using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Plugin_CalculateReservation
{
    public class Plugin_CalculateReservation : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        StringBuilder strMess = new StringBuilder();
        StringBuilder strMess2 = new StringBuilder();
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            ITracingService traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("Plugin_CalculateReservation");
            
            if (context.Depth > 4)
            {
                return;
            }
            
            if (context.MessageName == "Update")
            {
                traceService.Trace("vào case update");
                Entity target = (Entity)context.InputParameters["Target"];
                Entity quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                decimal bsd_detailamount = quote.Contains("bsd_detailamount") ? ((Money)quote["bsd_detailamount"]).Value : 0;
                decimal bsd_discountamount = quote.Contains("bsd_discountamount") ? ((Money)quote["bsd_discountamount"]).Value : 0;
                Entity up_quote = new Entity(quote.LogicalName, quote.Id);
                EntityReference PhaseLaunch = quote.Contains("bsd_phaseslaunchid") ? (EntityReference)quote["bsd_phaseslaunchid"] : null;
                if (PhaseLaunch != null)
                {
                    Entity enPhaseLaunch = service.Retrieve(PhaseLaunch.LogicalName, PhaseLaunch.Id, new ColumnSet(true));
                    if (enPhaseLaunch.Contains("bsd_depositamount"))
                    {
                        up_quote["bsd_depositfee"] = enPhaseLaunch["bsd_depositamount"];
                    }
                }
                if (quote.Contains("bsd_handovercondition"))
                {
                    EntityReference handoverRef = (EntityReference)quote["bsd_handovercondition"];

                    Entity enHandover = service.Retrieve(handoverRef.LogicalName,handoverRef.Id,new ColumnSet("bsd_amount"));

                    if (enHandover.Contains("bsd_amount"))
                    {
                        up_quote["bsd_packagesellingamount"] = enHandover["bsd_amount"];
                    }
                    decimal bsd_packagesellingamountValue = enHandover.Contains("bsd_amount") ? ((Money)enHandover["bsd_amount"]).Value : 0;
                    decimal bsd_totalamountlessfreight = bsd_detailamount - bsd_discountamount;
                    up_quote["bsd_totalamountlessfreight"] = new Money(bsd_totalamountlessfreight);
                    up_quote["bsd_packagesellingamount"] = new Money(bsd_packagesellingamountValue);
                    decimal bsd_totalamount = bsd_packagesellingamountValue + bsd_totalamountlessfreight;
                    up_quote["bsd_totalamount"] = new Money(bsd_totalamount);

                }
                service.Update(up_quote);
                

            }
        }
    }
}