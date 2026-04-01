using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Action_AreaAppendixy_PSGen
{
    public class Action_AreaAppendixy_PSGen : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService traceS = null;
        EntityReference target = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                target = (EntityReference)context.InputParameters["Target"];
                traceS = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceS.Trace($"start {target.Id}");
                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                traceS.Trace("2");
                deletePaymentSchemeDetail();
                if (enTarget.Contains("bsd_optionentry"))
                {
                    traceS.Trace("2");
                    EntityReference refOE = (EntityReference)enTarget["bsd_optionentry"];
                    genPaymentSchemeDetail(refOE);
                    decimal bsd_totalamountdifference = enTarget.Contains("bsd_totalamountdifference") ? ((Money)enTarget["bsd_totalamountdifference"]).Value : 0;
                    if (bsd_totalamountdifference > 0)
                    {
                        var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch top=""1"">
                          <entity name=""bsd_paymentschemedetail"">
                            <attribute name=""bsd_ordernumber"" />
                            <attribute name=""bsd_amountofthisphase"" />
                            <attribute name=""bsd_balance"" />
                            <filter>
                              <condition attribute=""bsd_areaappendix"" operator=""eq"" value=""{target.Id}"" />
                              <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                            </filter>
                            <order descending=""true"" attribute=""bsd_ordernumber"" />
                          </entity>
                        </fetch>";
                        EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        foreach (var item in rs.Entities)
                        {
                            decimal bsd_amountofthisphase = item.Contains("bsd_amountofthisphase") ? ((Money)item["bsd_amountofthisphase"]).Value : 0;
                            bsd_amountofthisphase += bsd_totalamountdifference;
                            decimal bsd_balance = item.Contains("bsd_balance") ? ((Money)item["bsd_balance"]).Value : 0;
                            bsd_balance += bsd_totalamountdifference;
                            Entity enUp = new Entity(item.LogicalName, item.Id);
                            enUp["bsd_amountofthisphase"] = new Money(bsd_amountofthisphase);
                            enUp["bsd_balance"] = new Money(bsd_balance);
                            service.Update(enUp);
                        }
                    }
                    else if (bsd_totalamountdifference < 0)
                    {
                        var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_paymentschemedetail"">
                            <attribute name=""bsd_ordernumber"" />
                            <attribute name=""bsd_amountofthisphase"" />
                            <attribute name=""bsd_balance"" />
                            <filter>
                              <condition attribute=""bsd_areaappendix"" operator=""eq"" value=""{target.Id}"" />
                              <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                            </filter>
                            <order descending=""true"" attribute=""bsd_ordernumber"" />
                          </entity>
                        </fetch>";
                        EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        foreach (var item in rs.Entities)
                        {
                            traceS.Trace("2");
                            decimal bsd_amountofthisphase = item.Contains("bsd_amountofthisphase") ? ((Money)item["bsd_amountofthisphase"]).Value : 0;
                            decimal bsd_balance = item.Contains("bsd_balance") ? ((Money)item["bsd_balance"]).Value : 0;
                            if (bsd_totalamountdifference <= bsd_balance)
                            {
                                bsd_balance -= bsd_totalamountdifference;
                                bsd_amountofthisphase -= bsd_totalamountdifference;
                                bsd_totalamountdifference = 0;
                                traceS.Trace("2");
                            }
                            else
                            {
                                bsd_amountofthisphase -= bsd_balance;
                                bsd_totalamountdifference -= bsd_balance;
                                Entity enNEW = new Entity("bsd_advancepayment");
                                enNEW["bsd_areaappendix"] = target;
                                enNEW["bsd_name"] = "Advance Payment";
                                enNEW["bsd_transactiondate"] = DateTime.Now;
                                enNEW["bsd_amount"] = new Money(bsd_totalamountdifference);
                                enNEW["bsd_remainingamount"] = new Money(bsd_totalamountdifference);
                                enNEW["bsd_remainingamountusd"] = new Money(bsd_totalamountdifference);
                                service.Create(enNEW);
                                bsd_balance = 0;
                                bsd_totalamountdifference = 0;
                                traceS.Trace("3");
                            }
                            Entity enUp = new Entity(item.LogicalName, item.Id);
                            enUp["bsd_amountofthisphase"] = new Money(bsd_amountofthisphase);
                            enUp["bsd_balance"] = new Money(bsd_balance);
                            service.Update(enUp);
                            if (bsd_totalamountdifference == 0) break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void deletePaymentSchemeDetail()
        {
            traceS.Trace("delete_PaymentSchemeDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""bsd_areaappendix"" operator=""eq"" value=""{target.Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (Entity item in rs.Entities)
                {
                    service.Delete(item.LogicalName, item.Id);
                }
            }
        }
        private void genPaymentSchemeDetail(EntityReference refOE)
        {
            traceS.Trace("genPaymentSchemeDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{refOE.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, target);
                }
            }
        }

        private void CreateNewFromItem(Entity item, EntityReference refTarget)
        {
            Entity it = new Entity(item.LogicalName);
            it = item;
            it.Attributes.Remove(item.LogicalName + "id");
            it.Attributes.Remove("ownerid");
            it.Attributes.Remove("bsd_quotation");
            it.Attributes.Remove("bsd_reservation");
            it.Attributes.Remove("bsd_reservationcontract");
            it.Attributes.Remove("bsd_optionentry");
            it["bsd_areaappendix"] = refTarget;
            it.Id = Guid.NewGuid();
            service.Create(it);
        }
    }
}