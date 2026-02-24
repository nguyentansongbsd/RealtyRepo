using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Plugin_OptionEntry_CalculateMoney
{
    public class Plugin_OptionEntry_CalculateMoney : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;
        IPluginExecutionContext context = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            trace.Trace($"{context.Depth}");
            if (context.Depth > (context.MessageName == "Create" ? 2 : 1))
                return;

            if (context.MessageName == "Create" || context.MessageName == "Update")
            {
                Entity target = (Entity)context.InputParameters["Target"];
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[]
                {
                    "bsd_detailamount", "bsd_discountcheck", "bsd_quoteid", "bsd_reservationcontract", "bsd_unitnumber", "bsd_taxcode", "bsd_handovercondition",
                    "bsd_pricelevel", "bsd_packagesellingamount", "bsd_freightamount", "bsd_totalamountlessfreight"
                }));

                Entity enUp = new Entity(target.LogicalName, target.Id);
                decimal unitPrice = GetUnitPrice(enTarget, target, ref enUp);

                if (context.MessageName == "Update")
                    CalcDiscount(enTarget, ref enUp, unitPrice);

                SetPrice(enTarget, ref enUp, unitPrice);

                service.Update(enUp);
            }
        }

        private void CalcDiscount(Entity enTarget, ref Entity enUp, decimal unitprice)
        {
            trace.Trace("CalcDiscount");

            delete_DiscountTransaction(enTarget.Id);
            if (!enTarget.Contains("bsd_discountcheck"))
            {
                enUp["bsd_discount"] = new Money(0);
                enUp["bsd_totalamountlessfreight"] = new Money(unitprice);
            }
            else
            {
                string[] strArray = enTarget["bsd_discountcheck"].ToString().Split(';');
                calculate_Discount_createDiscountTransaction(strArray, unitprice, enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice);
                enUp["bsd_discount"] = new Money(sumAmountDiscount);
                enUp["bsd_totalamountlessfreight"] = new Money(netSellingPrice);
            }
        }

        private void delete_DiscountTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_discounttransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_optionentry", ConditionOperator.Equal, idQuote));
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
                //    throw new InvalidPluginExecutionException(string.Format("Discount '{0}' dose not exist or deleted.", pro["bsd_name"]));
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
                rsv["bsd_optionentry"] = enTarget.ToEntityReference();
                service.Create(rsv);
                no++;
            }
        }

        private void SetPrice(Entity enOE, ref Entity enUp, decimal bsd_detailamount)
        {
            decimal bsd_totalamountlessfreight = 0;
            if (context.MessageName == "Create")
                bsd_totalamountlessfreight = GetMoney(enOE, "bsd_totalamountlessfreight");
            else
                bsd_totalamountlessfreight = GetMoney(enUp, "bsd_totalamountlessfreight");
            decimal bsd_freightamount = 0;

            if (!enOE.Contains("bsd_quoteid") && !enOE.Contains("bsd_reservationcontract"))  // từ sp
            {
                trace.Trace("từ sp");

                decimal bsd_packagesellingamount = GetPackageSellingAmount(enOE, bsd_detailamount);
                enUp["bsd_packagesellingamount"] = new Money(bsd_packagesellingamount);

                bsd_freightamount = GetFreightAmount(enOE, bsd_totalamountlessfreight);
                enUp["bsd_freightamount"] = new Money(bsd_freightamount);

                SetTotalTaxAndTotalAmount(enOE, bsd_totalamountlessfreight, bsd_freightamount, ref enUp);
            }
            else  // từ convert
            {
                if (context.MessageName == "Update")
                {
                    trace.Trace("từ convert");
                    bsd_freightamount = GetMoney(enOE, "bsd_freightamount");

                    SetTotalTaxAndTotalAmount(enOE, bsd_totalamountlessfreight, bsd_freightamount, ref enUp);
                }
            }
        }

        private decimal GetPackageSellingAmount(Entity enOE, decimal bsd_detailamount)
        {
            trace.Trace("GetPackageSellingAmount");

            decimal bsd_packagesellingamount = 0;
            if (!enOE.Contains("bsd_handovercondition"))
                return bsd_packagesellingamount;

            EntityReference refHandover = (EntityReference)enOE["bsd_handovercondition"];
            Entity enHandover = service.Retrieve(refHandover.LogicalName, refHandover.Id, new ColumnSet(new string[] { "bsd_method", "bsd_amount", "bsd_percent" }));
            int bsd_method = enHandover.Contains("bsd_method") ? ((OptionSetValue)enHandover["bsd_method"]).Value : -99;

            if (bsd_method == 100000001)    //Amount
            {
                bsd_packagesellingamount = GetMoney(enHandover, "bsd_amount");
            }
            else if (bsd_method == 100000002)   //Percent (%)
            {
                decimal bsd_percent = enHandover.Contains("bsd_percent") ? (decimal)enHandover["bsd_percent"] / 100 : 0;
                bsd_packagesellingamount = bsd_detailamount * bsd_percent;
            }

            return bsd_packagesellingamount;
        }

        private decimal GetFreightAmount(Entity enOE, decimal bsd_totalamountlessfreight)
        {
            trace.Trace("GetFreightAmount");

            EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];
            Entity enUnit = service.Retrieve(refUnit.LogicalName, refUnit.Id, new ColumnSet(new string[] { "bsd_maintenancefeespercent" }));
            decimal bsd_maintenancefeespercent = enUnit.Contains("bsd_maintenancefeespercent") ? (decimal)enUnit["bsd_maintenancefeespercent"] / 100 : 0;
            return bsd_maintenancefeespercent * bsd_totalamountlessfreight;
        }

        private void SetTotalTaxAndTotalAmount(Entity enOE, decimal bsd_totalamountlessfreight, decimal bsd_freightamount, ref Entity enUp)
        {
            trace.Trace("SetTotalTaxAndTotalAmount");

            decimal percentTax = 0;
            if (enOE.Contains("bsd_taxcode"))
            {
                EntityReference refTax = (EntityReference)enOE["bsd_taxcode"];
                Entity enTax = service.Retrieve(refTax.LogicalName, refTax.Id, new ColumnSet(new string[] { "bsd_value" }));
                percentTax = enTax.Contains("bsd_value") ? (decimal)enTax["bsd_value"] / 100 : 0;
            }
            decimal bsd_totaltax = bsd_totalamountlessfreight * percentTax;
            enUp["bsd_totaltax"] = new Money(bsd_totaltax);
            decimal bsd_totalamountlessfreightaftervat = bsd_totalamountlessfreight + bsd_totaltax;
            enUp["bsd_totalamountlessfreightaftervat"] = new Money(bsd_totalamountlessfreightaftervat);
            enUp["bsd_totalamount"] = new Money(bsd_totalamountlessfreightaftervat + bsd_freightamount);
        }

        private decimal GetUnitPrice(Entity enTarget, Entity target, ref Entity enUp)
        {
            trace.Trace("GetUnitPrice");
            decimal unitPrice = 0;

            if (context.MessageName == "Update" && target.Contains("bsd_pricelevel") && !enTarget.Contains("bsd_quoteid") && !enTarget.Contains("bsd_reservationcontract"))
            {
                unitPrice = GetNewListedPrice(enTarget);
                enUp["bsd_detailamount"] = new Money(unitPrice);
            }
            else
                unitPrice = GetMoney(enTarget, "bsd_detailamount");

            return unitPrice;
        }

        private decimal GetNewListedPrice(Entity enOE)
        {
            trace.Trace("GetNewListedPrice");
            EntityReference refPriceList = (EntityReference)enOE["bsd_pricelevel"];
            EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_productpricelevel"">
                <attribute name=""bsd_price"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{refPriceList.Id}"" />
                  <condition attribute=""bsd_product"" operator=""eq"" value=""{refUnit.Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count == 1)
            {
                return GetMoney(rs.Entities[0], "bsd_price");
            }

            return 0;
        }

        private decimal GetMoney(Entity e, string field)
        {
            return e.Contains(field) ? ((Money)e[field]).Value : 0;
        }
    }
}