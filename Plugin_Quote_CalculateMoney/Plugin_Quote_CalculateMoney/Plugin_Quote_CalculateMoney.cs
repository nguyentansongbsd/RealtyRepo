using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Plugin_Quote_CalculateMoney
{
    public class Plugin_Quote_CalculateMoney : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;
        IPluginExecutionContext context = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.Depth > 1) return;
            if (context.MessageName == "Create" || context.MessageName == "Update")
            {
                Entity target = (Entity)context.InputParameters["Target"];
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[]
                {
                    "bsd_detailamount",
                    "bsd_discountcheck"
                }));

                decimal unitprice = enTarget.Contains("bsd_detailamount") ? ((Money)enTarget["bsd_detailamount"]).Value : 0;
                delete_DiscountTransaction(target.Id);
                if (!enTarget.Contains("bsd_discountcheck"))
                {
                    Entity enUp = new Entity(target.LogicalName, target.Id);
                    enUp["bsd_discountamount"] = new Money(0);
                    enUp["bsd_totalamountlessfreight"] = new Money(unitprice);
                    service.Update(enUp);
                }
                else
                {
                    string[] strArray = enTarget["bsd_discountcheck"].ToString().Split(';');
                    calculate_Discount_createDiscountTransaction(strArray, unitprice, enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice);
                    Entity enUp = new Entity(target.LogicalName, target.Id);
                    enUp["bsd_discountamount"] = new Money(sumAmountDiscount);
                    enUp["bsd_totalamountlessfreight"] = new Money(netSellingPrice);
                    service.Update(enUp);
                }
            }
        }
        private void delete_DiscountTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_discounttransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_quote", ConditionOperator.Equal, idQuote));
            EntityCollection entityCollection1 = this.service.RetrieveMultiple(queryExpression);
            Dictionary<Guid, string> dictionary = new Dictionary<Guid, string>();
            foreach (Entity entity in entityCollection1.Entities)
                service.Delete(entity.LogicalName, entity.Id);
        }
        private void calculate_Discount_createDiscountTransaction(string[] strArray, decimal unitprice, Entity enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice)
        {
            sumAmountDiscount = 0;
            netSellingPrice = unitprice;
            //List<Guid> strArrayAmount = new List<Guid>();
            //List<Guid> strArrayPercent = new List<Guid>();
            List<Guid> strArraySum = new List<Guid>();
            foreach (string input in strArray)
            {
                Guid guid = Guid.Parse(input);
                strArraySum.Add(guid);
                //Entity pro = service.Retrieve("bsd_discount", guid, new ColumnSet(new string[1] { "bsd_method" }));

                //if (pro == null)
                //    throw new InvalidPluginExecutionException(string.Format("Discount '{0}' dose not exist or deleted.", pro["bsd_name"]) + MessageProvider.GetMessage(service, context, "check_percent_ins"));
                //if (!pro.Contains("bsd_method"))
                //    throw new InvalidPluginExecutionException(string.Format("Please provide method for discount '{0}'!", pro["bsd_name"]));
                //int num = ((OptionSetValue)pro["bsd_method"]).Value;
                //if (num == 100000001)//percent
                //{
                //    strArrayPercent.Add(guid);
                //}
                //else
                //{
                //    strArrayAmount.Add(guid);
                //}
            }
            //strArraySum = strArrayAmount.Concat(strArrayPercent).Distinct().ToList();
            int no = 1;
            foreach (Guid guid in strArraySum)
            {
                Entity pro = service.Retrieve("bsd_discount", guid, new ColumnSet(new string[5]
                    {
                        "bsd_name",
                        "bsd_method",
                        "bsd_amount",
                        "bsd_percentage",
                        "bsd_type"
                    }));
                if (pro == null)
                    throw new InvalidPluginExecutionException(string.Format("Discount '{0}' dose not exist or deleted.", pro["bsd_name"]));
                if (!pro.Contains("bsd_method"))
                    throw new InvalidPluginExecutionException(string.Format("Please provide method for discount '{0}'!", pro["bsd_name"]));
                int type = ((OptionSetValue)pro["bsd_type"]).Value;
                int num = ((OptionSetValue)pro["bsd_method"]).Value;
                Entity rsv = new Entity("bsd_discounttransaction");
                rsv["bsd_no"] = no;
                trace.Trace("name discount: " + pro["bsd_name"]);
                if (num == 100000001)//percent
                {
                    trace.Trace("percent");
                    if (!pro.Contains("bsd_percentage"))
                        throw new InvalidPluginExecutionException(string.Format("Please provide discount percent for discount'{0}'", pro["bsd_name"]));
                    decimal bsd_percentage = (decimal)pro["bsd_percentage"];
                    if (!pro.Contains("bsd_type"))
                        throw new InvalidPluginExecutionException(string.Format("Please provide type for discount '{0}'!", pro["bsd_name"]));
                    rsv["bsd_name"] = pro["bsd_name"];
                    rsv["bsd_discountpercent"] = bsd_percentage;
                    trace.Trace("bsd_percentage: " + bsd_percentage);
                    trace.Trace("type: " + type);
                    decimal amountDiscount = Math.Round(bsd_percentage * (type == 100000000 ? netSellingPrice : unitprice) / 100, MidpointRounding.AwayFromZero);
                    trace.Trace("amountDiscount: " + amountDiscount);
                    rsv["bsd_totaldiscountamount"] = new Money(amountDiscount);
                    sumAmountDiscount += amountDiscount;
                    netSellingPrice -= amountDiscount;
                    trace.Trace("sumAmountDiscount: " + sumAmountDiscount);
                    trace.Trace("netSellingPrice: " + netSellingPrice);
                }
                else
                {
                    trace.Trace("amount");
                    if (!pro.Contains("bsd_amount"))
                        throw new InvalidPluginExecutionException(string.Format("Please provide discount amount for discount'{0}'", pro["bsd_name"]));
                    decimal bsd_amount = ((Money)pro["bsd_amount"]).Value;
                    trace.Trace("bsd_amount: " + bsd_amount);
                    rsv["bsd_name"] = pro["bsd_name"];
                    rsv["bsd_discountamount"] = new Money(bsd_amount);
                    rsv["bsd_totaldiscountamount"] = new Money(bsd_amount);
                    sumAmountDiscount += bsd_amount;
                    netSellingPrice -= bsd_amount;
                    trace.Trace("sumAmountDiscount: " + sumAmountDiscount);
                    trace.Trace("netSellingPrice: " + netSellingPrice);
                }
                rsv["bsd_discount"] = pro.ToEntityReference();
                rsv["bsd_quote"] = enTarget.ToEntityReference();
                service.Create(rsv);
                no++;
            }
        }
    }
}