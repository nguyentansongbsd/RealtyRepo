using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_ConstructionMilestone_Delete
{
    public class Plugin_ConstructionMilestone_Delete : IPlugin
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

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch top=""1"">
                  <entity name=""bsd_constructionmilestone"">
                    <attribute name=""bsd_constructionmilestoneid"" />
                    <attribute name=""bsd_name"" />
                    <filter>
                      <condition attribute=""bsd_constructionmilestoneid"" operator=""eq"" value=""{target.Id}"" />
                    </filter>
                    <filter type=""or"">
                      <condition entityname=""bsd_paymentschemedetail"" attribute=""bsd_paymentschemedetailid"" operator=""not-null"" />
                      <condition entityname=""bsd_paymentschemedetailmaster"" attribute=""bsd_paymentschemedetailmasterid"" operator=""not-null"" />
                    </filter>
                    <link-entity name=""bsd_paymentschemedetail"" from=""bsd_constructionmilestone"" to=""bsd_constructionmilestoneid"" link-type=""outer"" />
                    <link-entity name=""bsd_paymentschemedetailmaster"" from=""bsd_constructionmilestone"" to=""bsd_constructionmilestoneid"" link-type=""outer"" />
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "delete_milestone"));
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