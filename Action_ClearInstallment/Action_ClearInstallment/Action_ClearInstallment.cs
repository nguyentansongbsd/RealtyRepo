using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_ClearInstallment
{
    public class Action_ClearInstallment : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("start");
            if (context.Depth > 2) return;

            string id = (string)context.InputParameters["id"];
            string logicalName = (string)context.InputParameters["logicalNameLookup"];

            traceService.Trace("bắt đầu xóa");
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""{logicalName}"" operator=""eq"" value=""{id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    service.Delete(item.LogicalName, item.Id);
                }
            }
            traceService.Trace("done");
        }
    }
}
