using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Plugin_Reservation_CalculateMoney
{
    public class Plugin_Reservation_CalculateMoney : IPlugin
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
                if (context.MessageName == "Update" && target.Contains("bsd_customerid"))
                {
                    EntityReference customer = (EntityReference)target["bsd_customerid"];
                    List<string> missingFields = new List<string>();

                    // Define field mapping
                    Dictionary<string, string> fieldMap = new Dictionary<string, string>();

                    if (customer.LogicalName == "contact")
                    {
                        fieldMap = new Dictionary<string, string>
                            {
                                {"bsd_localization", "Nationality"},
                                {"birthdate", "Birthday"},
                                {"bsd_identitycardnumber", "Identity Card Number (ID)"},
                                {"bsd_country", "Country"},
                                {"bsd_province", "Province"},
                                {"bsd_contactaddress", "Contact Address (VN)"},
                                {"bsd_permanentcountry", "Permanent Country"},
                                {"bsd_permanentprovince", "Permanent Province"},
                                {"bsd_permanentaddress1", "Permanent Address (VN)"}
                            };
                    }
                    else if (customer.LogicalName == "account")
                    {
                        fieldMap = new Dictionary<string, string>
                            {
                                {"bsd_localization", "Nationality"},
                                {"bsd_registrationcode", "Registration Code"},
                                {"bsd_nation", "Country"},
                                {"bsd_province", "Province"},
                                {"bsd_addressvn", "Address (VN)"},
                                {"bsd_permanentcountry", "Permanent Country"},
                                {"bsd_permanentprovince", "Permanent Province"},
                                {"bsd_permanentaddress1", "Permanent Address (VN)"}
                            };
                    }

                    // Retrieve only needed columns
                    Entity customerEntity = service.Retrieve(
                        customer.LogicalName,
                        customer.Id,
                        new ColumnSet(fieldMap.Keys.ToArray())
                    );

                    // Validate null / missing
                    foreach (var field in fieldMap)
                    {
                        if (!customerEntity.Contains(field.Key) || customerEntity[field.Key] == null)
                        {
                            missingFields.Add(field.Value);
                        }
                    }

                    // Throw error if needed
                    if (missingFields.Count > 0)
                    {
                        throw new InvalidPluginExecutionException(
                            "Please fill in the missing customer information below:\r\n ["
                            + string.Join("\r\n| ", missingFields) + "]"
                        );
                    }
                }
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[]
                {
                    "bsd_detailamount",
                    "bsd_promotioncheck",
                    "bsd_discountcheck"
                }));

                decimal unitprice = enTarget.Contains("bsd_detailamount") ? ((Money)enTarget["bsd_detailamount"]).Value : 0;
                delete_DiscountTransaction(target.Id);
                Entity enUp = new Entity(target.LogicalName, target.Id);
                decimal totalamountlessfreight = unitprice;
                if (!enTarget.Contains("bsd_discountcheck"))
                {
                    enUp["bsd_discountamount"] = new Money(0);
                    enUp["bsd_totalamountlessfreight"] = new Money(unitprice);
                }
                else
                {
                    string[] strArray = enTarget["bsd_discountcheck"].ToString().Split(';');
                    calculate_Discount_createDiscountTransaction(strArray, unitprice, enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice);
                    enUp["bsd_discountamount"] = new Money(sumAmountDiscount);
                    enUp["bsd_totalamountlessfreight"] = new Money(netSellingPrice);
                }
                delete_PromotionTransaction(target.Id);
                if (!enTarget.Contains("bsd_promotioncheck"))
                {
                    enUp["bsd_promotion"] = new Money(0);
                    enUp["bsd_totalamountlessfreight"] = new Money(totalamountlessfreight);
                }
                else
                {
                    string[] strArray = enTarget["bsd_promotioncheck"].ToString().Split(';');
                    calculate_Promotion_createPromotionTransaction(strArray, unitprice, enTarget, out decimal sumAmountPromotion);
                    totalamountlessfreight -= sumAmountPromotion;
                    enUp["bsd_promotion"] = new Money(sumAmountPromotion);
                    enUp["bsd_totalamountlessfreight"] = new Money(totalamountlessfreight);
                }
                service.Update(enUp);
            }
        }
        private void delete_DiscountTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_discounttransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_reservationcontract", ConditionOperator.Equal, idQuote));
            EntityCollection entityCollection1 = this.service.RetrieveMultiple(queryExpression);
            Dictionary<Guid, string> dictionary = new Dictionary<Guid, string>();
            foreach (Entity entity in entityCollection1.Entities)
                service.Delete(entity.LogicalName, entity.Id);
        }
        private void calculate_Discount_createDiscountTransaction(string[] strArray, decimal unitprice, Entity enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice)
        {
            sumAmountDiscount = 0;
            netSellingPrice = unitprice;
            List<Guid> strArraySum = new List<Guid>();
            foreach (string input in strArray)
            {
                Guid guid = Guid.Parse(input);
                strArraySum.Add(guid);
            }
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
                rsv["bsd_reservationcontract"] = enTarget.ToEntityReference();
                service.Create(rsv);
                no++;
            }
        }
        private void delete_PromotionTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_promotiontransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_ra", ConditionOperator.Equal, idQuote));
            EntityCollection entityCollection1 = this.service.RetrieveMultiple(queryExpression);
            Dictionary<Guid, string> dictionary = new Dictionary<Guid, string>();
            foreach (Entity entity in entityCollection1.Entities)
                service.Delete(entity.LogicalName, entity.Id);
        }
        private void calculate_Promotion_createPromotionTransaction(string[] strArray, decimal unitprice, Entity enTarget, out decimal sumAmountPromotion)
        {
            sumAmountPromotion = 0;
            List<Guid> strArraySum = new List<Guid>();
            foreach (string input in strArray)
            {
                Guid guid = Guid.Parse(input);
                strArraySum.Add(guid);
            }
            foreach (Guid guid in strArraySum)
            {
                trace.Trace("vào for promotion");
                Entity pro = service.Retrieve("bsd_promotion", guid, new ColumnSet(new string[5]
                    {
                        "bsd_name",
                        "bsd_method",
                        "bsd_amount",
                        "bsd_percent",
                        "bsd_contractdeduction"
                    }));
                if (pro == null)
                    throw new InvalidPluginExecutionException(string.Format("Promotion '{0}' dose not exist or deleted.", pro["bsd_name"]));
                if (!pro.Contains("bsd_method"))
                    throw new InvalidPluginExecutionException(string.Format("Please provide method for promotion '{0}'!", pro["bsd_name"]));
                bool bsd_contractdeduction = pro.Contains("bsd_contractdeduction") ? (bool)pro["bsd_contractdeduction"] : false;
                if (bsd_contractdeduction)
                {
                    int num = ((OptionSetValue)pro["bsd_method"]).Value;
                    Entity rsv = new Entity("bsd_promotiontransaction");
                    trace.Trace("name discount: " + pro["bsd_name"]);
                    if (num == 100000001)//percent
                    {
                        trace.Trace("percent");
                        if (!pro.Contains("bsd_percent"))
                            throw new InvalidPluginExecutionException(string.Format("Please provide promotion percent for promotion'{0}'", pro["bsd_name"]));
                        decimal bsd_percentage = (decimal)pro["bsd_percent"];
                        rsv["bsd_name"] = pro["bsd_name"];
                        rsv["bsd_promotionpercent"] = bsd_percentage;
                        trace.Trace("bsd_promotionpercent: " + bsd_percentage);
                        decimal amountDiscount = Math.Round(bsd_percentage * unitprice / 100, MidpointRounding.AwayFromZero);
                        trace.Trace("amountDiscount: " + amountDiscount);
                        rsv["bsd_totalpromotionamount"] = new Money(amountDiscount);
                        sumAmountPromotion += amountDiscount;
                        trace.Trace("sumAmountPromotion: " + sumAmountPromotion);
                    }
                    else
                    {
                        trace.Trace("amount");
                        if (!pro.Contains("bsd_amount"))
                            throw new InvalidPluginExecutionException(string.Format("Please provide promotion amount for promotion'{0}'", pro["bsd_name"]));
                        decimal bsd_amount = ((Money)pro["bsd_amount"]).Value;
                        trace.Trace("bsd_amount: " + bsd_amount);
                        rsv["bsd_name"] = pro["bsd_name"];
                        rsv["bsd_promotionamount"] = new Money(bsd_amount);
                        rsv["bsd_totalpromotionamount"] = new Money(bsd_amount);
                        sumAmountPromotion += bsd_amount;
                        trace.Trace("sumAmountPromotion: " + sumAmountPromotion);
                    }
                    trace.Trace("vào promotion");
                    rsv["bsd_promotion"] = pro.ToEntityReference();
                    trace.Trace("map promotion");
                    rsv["bsd_ra"] = enTarget.ToEntityReference();
                    trace.Trace("map quote");
                    service.Create(rsv);
                }
            }
        }
    }
}