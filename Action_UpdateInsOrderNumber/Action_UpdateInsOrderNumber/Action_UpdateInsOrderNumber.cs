using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_UpdateInsOrderNumber
{
    public class Action_UpdateInsOrderNumber : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 1) return;

                string id = (string)context.InputParameters["id"];
                if (string.IsNullOrEmpty(id))
                    return;

                int priceType = (int)context.InputParameters["priceType"];
                if (priceType == 0)
                    return;

                var query = new QueryExpression("bsd_paymentschemedetailmaster");
                query.ColumnSet.AddColumns("bsd_ordernumber", "bsd_name");
                query.Criteria.AddCondition("bsd_paymentscheme", ConditionOperator.Equal, id);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("bsd_pricetype", ConditionOperator.Equal, priceType);

                query.AddOrder("bsd_ordernumber", OrderType.Ascending);
                EntityCollection rs = service.RetrieveMultiple(query);
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    int expectedOrder = 1;
                    foreach (var item in rs.Entities)
                    {
                        int currentOrder = item.Contains("bsd_ordernumber") ? (int)item["bsd_ordernumber"] : 0;
                        if (currentOrder != expectedOrder)
                        {
                            Entity upIns = new Entity("bsd_paymentschemedetailmaster", item.Id);
                            upIns["bsd_ordernumber"] = expectedOrder;
                            upIns["bsd_name"] = $"Đợt {expectedOrder}";
                            service.Update(upIns);
                        }

                        expectedOrder++;
                    }

                    traceService.Trace("done");
                }
                else
                    throw new InvalidPluginExecutionException("Không có đợt nào tồn tại. Xin vui lòng kiểm tra lại");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}