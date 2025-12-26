using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_GenInsByTimes
{
    public class Action_GenInsByTimes : IPlugin
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
            if (context.Depth > 1) return;

            string id = (string)context.InputParameters["id"];
            if (string.IsNullOrEmpty(id))
                return;

            int priceType = (int)context.InputParameters["priceType"];

            var query = new QueryExpression("bsd_paymentschemedetailmaster");
            query.ColumnSet.AddColumns(
                "bsd_paymentschemedetailmasterid",
                "bsd_name",
                "bsd_nextperiodtype",
                "bsd_numberofnextmonth",
                "bsd_startfrominstallment",
                "bsd_numberofnextdays",
                "bsd_number");
            query.Criteria.AddCondition("bsd_paymentscheme", ConditionOperator.Equal, id);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("bsd_pricetype", ConditionOperator.Equal, priceType);
            query.Criteria.AddCondition("bsd_typepayment", ConditionOperator.Equal, 2);
            query.Criteria.AddCondition("bsd_number", ConditionOperator.NotNull);
            EntityCollection rs = service.RetrieveMultiple(query);
            if (rs == null || rs.Entities == null || rs.Entities.Count == 0)
                throw new InvalidPluginExecutionException("Vui lòng tạo đợt thanh toán có 'Payment Type' là 'Times'.");

            query = new QueryExpression("bsd_paymentschemedetailmaster");
            query.ColumnSet.AllColumns = true;
            query.Criteria.AddCondition("bsd_paymentscheme", ConditionOperator.Equal, id);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("bsd_pricetype", ConditionOperator.Equal, priceType);
            query.AddOrder("bsd_ordernumber", OrderType.Ascending);
            rs = service.RetrieveMultiple(query);
            traceService.Trace($"rs {rs.Entities.Count}");
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                int order = 0;
                bool flag = false;
                List<Entity> listCreate = new List<Entity>();
                for (int i = 0; i < rs.Entities.Count; i++)
                {
                    Entity item = rs.Entities[i];
                    order++;
                    traceService.Trace($"ins {item.Id} || {order}");
                    if (!(item.Contains("bsd_typepayment") && ((OptionSetValue)item["bsd_typepayment"]).Value == 2
                            && item.Contains("bsd_number") && (int)item["bsd_number"] > 0))
                    {
                        if (flag)
                        {
                            traceService.Trace($"cập nhật lại order {item.Id} || {order}");
                            Entity upIns = new Entity(item.LogicalName, item.Id);
                            upIns["bsd_ordernumber"] = order;
                            upIns["bsd_name"] = $"Đợt {order}";
                            service.Update(upIns);
                        }
                        continue;
                    }

                    flag = true;
                    Guid idNewIns = item.Id;
                    int number = (int)item["bsd_number"];
                    traceService.Trace($"gen {idNewIns} || {number}");
                    for (int j = 0; j < number - 1; j++)
                    {
                        order++;
                        Entity newIns = new Entity(item.LogicalName);

                        newIns["bsd_ordernumber"] = order;
                        newIns["bsd_name"] = $"Đợt {order}";
                        newIns["bsd_startfrominstallment"] = new EntityReference(item.LogicalName, idNewIns);
                        newIns["bsd_number"] = 0;

                        newIns["bsd_project"] = item.Contains("bsd_project") ? item["bsd_project"] : null;
                        newIns["bsd_paymentscheme"] = item.Contains("bsd_paymentscheme") ? item["bsd_paymentscheme"] : null;
                        newIns["bsd_calendartype"] = item.Contains("bsd_calendartype") ? item["bsd_calendartype"] : null;
                        newIns["bsd_pricetype"] = item.Contains("bsd_pricetype") ? item["bsd_pricetype"] : null;
                        newIns["bsd_duedatecalculatingmethod"] = item.Contains("bsd_duedatecalculatingmethod") ? item["bsd_duedatecalculatingmethod"] : null;
                        newIns["bsd_calculationmethod"] = item.Contains("bsd_calculationmethod") ? item["bsd_calculationmethod"] : null;
                        newIns["bsd_nextperiodtype"] = item.Contains("bsd_nextperiodtype") ? item["bsd_nextperiodtype"] : null;
                        newIns["bsd_numberofnextdays"] = item.Contains("bsd_numberofnextdays") ? item["bsd_numberofnextdays"] : null;
                        newIns["bsd_numberofnextmonth"] = item.Contains("bsd_numberofnextmonth") ? item["bsd_numberofnextmonth"] : null;
                        newIns["bsd_typepayment"] = item.Contains("bsd_typepayment") ? item["bsd_typepayment"] : null;
                        newIns["bsd_amount"] = item.Contains("bsd_amount") ? item["bsd_amount"] : null;
                        newIns["bsd_amountpercent"] = item.Contains("bsd_amountpercent") ? item["bsd_amountpercent"] : null;
                        newIns["bsd_official"] = true;
                        newIns["bsd_gopdot"] = true;
                        idNewIns = Guid.NewGuid();
                        newIns.Id = idNewIns;
                        listCreate.Add(newIns);
                    }
                }

                traceService.Trace($"listCreate {listCreate.Count}");
                foreach (var item in listCreate)
                {
                    service.Create(item);
                }
            }
        }
    }
}