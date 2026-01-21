using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Quote_Approve
{
    public class Plugin_Quote_Approve : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity target = context.InputParameters["Target"] as Entity;

            Entity quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet("statuscode", "bsd_totalamountpaid", "bsd_minimumdeposit", "bsd_approvereason"));
            int status = ((OptionSetValue)quote["statuscode"]).Value;
            if (status == 667980001)
            {
                trace.Trace("Vào Plugin_Quote_Approve");
                decimal totalamount = quote.Contains("bsd_totalamountpaid") ? ((Money)quote["bsd_totalamountpaid"]).Value : 0;
                decimal bsd_minimumdeposit = quote.Contains("bsd_minimumdeposit") ? ((Money)quote["bsd_minimumdeposit"]).Value : 0;
                if (totalamount < bsd_minimumdeposit)
                {
                    if(!quote.Contains("bsd_approvereason"))
                    {
                        trace.Trace("End Plugin_Quote_Approve");
                        throw new InvalidPluginExecutionException("\nPlease pay the minimum required deposit amount.");
                    }
                }
            }
        }
    }
}
