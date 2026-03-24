using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Plugin_Quotation_CalculateMoney
{
    public class Plugin_Quotation_CalculateMoney : IPlugin
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
                    "bsd_detailamount", "bsd_discountcheck", "bsd_promotioncheck", "bsd_unitnumber", "bsd_taxcode", "bsd_handovercondition", "bsd_maintenancefeespercent",
                    "bsd_pricelevel", "bsd_packagesellingamount", "bsd_freightamount", "bsd_totalamountlessfreight", "bsd_landvaluededuction", "bsd_vat"
                }));

                Entity enUp = new Entity(target.LogicalName, target.Id);
                decimal unitPrice = GetUnitPrice(enTarget, target, ref enUp);
                decimal bsd_totalamountlessfreight = unitPrice;

                if (context.MessageName == "Update")
                    CalcDiscount(enTarget, ref enUp, unitPrice, ref bsd_totalamountlessfreight);

                SetPrice(enTarget, ref enUp, unitPrice, bsd_totalamountlessfreight);

                service.Update(enUp);
            }
        }

        private void CalcDiscount(Entity enTarget, ref Entity enUp, decimal unitPrice, ref decimal bsd_totalamountlessfreight)
        {
            trace.Trace("CalcDiscount");

            delete_DiscountTransaction(enTarget.Id);
            if (!enTarget.Contains("bsd_discountcheck"))
            {
                enUp["bsd_discountamount"] = new Money(0);
                bsd_totalamountlessfreight = unitPrice;
            }
            else
            {
                string[] strArray = enTarget["bsd_discountcheck"].ToString().Split(';');
                calculate_Discount_createDiscountTransaction(strArray, unitPrice, enTarget, out decimal sumAmountDiscount, out decimal netSellingPrice);
                enUp["bsd_discountamount"] = new Money(sumAmountDiscount);
                bsd_totalamountlessfreight = netSellingPrice;
            }
            delete_PromotionTransaction(enTarget.Id);
            if (!enTarget.Contains("bsd_promotioncheck"))
            {
                enUp["bsd_promotion"] = new Money(0);
            }
            else
            {
                string[] strArray = enTarget["bsd_promotioncheck"].ToString().Split(';');
                calculate_Promotion_createPromotionTransaction(strArray, unitPrice, enTarget, out decimal sumAmountPromotion);
                bsd_totalamountlessfreight -= sumAmountPromotion;
                enUp["bsd_promotion"] = new Money(sumAmountPromotion);
            }
        }

        private void delete_DiscountTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_discounttransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_quotation", ConditionOperator.Equal, idQuote));
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
                trace.Trace("vào for");
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
                trace.Trace("vào dis");
                rsv["bsd_discount"] = pro.ToEntityReference();
                trace.Trace("map dis");
                rsv["bsd_quotation"] = enTarget.ToEntityReference();
                trace.Trace("map quote");
                service.Create(rsv);
                no++;
            }
        }
        private void delete_PromotionTransaction(Guid idQuote)
        {
            QueryExpression queryExpression = new QueryExpression("bsd_promotiontransaction");
            queryExpression.ColumnSet = new ColumnSet(new string[0]);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition(new ConditionExpression("bsd_quotation", ConditionOperator.Equal, idQuote));
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
                    rsv["bsd_quotation"] = enTarget.ToEntityReference();
                    trace.Trace("map quote");
                    service.Create(rsv);
                }
            }
        }

        private void SetPrice(Entity enQuotation, ref Entity enUp, decimal bsd_detailamount, decimal bsd_totalamountlessfreight)
        {
            trace.Trace("SetPrice");

            decimal bsd_packagesellingamount = GetPackageSellingAmount(enQuotation, bsd_detailamount);
            enUp["bsd_packagesellingamount"] = new Money(bsd_packagesellingamount);

            bsd_totalamountlessfreight += bsd_packagesellingamount;
            enUp["bsd_totalamountlessfreight"] = new Money(bsd_totalamountlessfreight);

            decimal bsd_freightamount = GetFreightAmount(enQuotation, bsd_totalamountlessfreight);
            enUp["bsd_freightamount"] = new Money(bsd_freightamount);

            decimal bsd_vat = GetVAT(enQuotation, bsd_totalamountlessfreight);
            enUp["bsd_vat"] = new Money(bsd_vat);

            decimal bsd_totalamountlessfreightaftervat = bsd_totalamountlessfreight + bsd_vat;
            enUp["bsd_totalamountlessfreightaftervat"] = new Money(bsd_totalamountlessfreightaftervat);
            enUp["bsd_totalamount"] = new Money(bsd_totalamountlessfreightaftervat + bsd_freightamount);
        }

        private decimal GetPackageSellingAmount(Entity enQuotation, decimal bsd_detailamount)
        {
            trace.Trace("GetPackageSellingAmount");

            decimal bsd_packagesellingamount = 0;
            if (!enQuotation.Contains("bsd_handovercondition"))
                return bsd_packagesellingamount;

            EntityReference refHandover = (EntityReference)enQuotation["bsd_handovercondition"];
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

        private decimal GetFreightAmount(Entity enQuotation, decimal bsd_totalamountlessfreight)
        {
            trace.Trace("GetFreightAmount");

            decimal bsd_maintenancefeespercent = enQuotation.Contains("bsd_maintenancefeespercent") ? (decimal)enQuotation["bsd_maintenancefeespercent"] / 100 : 0;
            return bsd_maintenancefeespercent * bsd_totalamountlessfreight;
        }

        private decimal GetVAT(Entity enQuotation, decimal bsd_totalamountlessfreight)
        {
            trace.Trace("GetVAT");

            decimal percentTax = 0;
            if (enQuotation.Contains("bsd_taxcode"))
            {
                EntityReference refTax = (EntityReference)enQuotation["bsd_taxcode"];
                Entity enTax = service.Retrieve(refTax.LogicalName, refTax.Id, new ColumnSet(new string[] { "bsd_value" }));
                percentTax = enTax.Contains("bsd_value") ? (decimal)enTax["bsd_value"] / 100 : 0;
            }
            decimal bsd_landvaluededuction = GetMoney(enQuotation, "bsd_landvaluededuction");
            decimal bsd_vat = (bsd_totalamountlessfreight - bsd_landvaluededuction) * percentTax;
            return bsd_vat;
        }

        private decimal GetUnitPrice(Entity enTarget, Entity target, ref Entity enUp)
        {
            trace.Trace("GetUnitPrice");
            decimal unitPrice = 0;

            if (context.MessageName == "Update" && target.Contains("bsd_pricelevel"))
            {
                unitPrice = GetNewListedPrice(enTarget);
                enUp["bsd_detailamount"] = new Money(unitPrice);
            }
            else
                unitPrice = GetMoney(enTarget, "bsd_detailamount");

            return unitPrice;
        }

        private decimal GetNewListedPrice(Entity enQuotation)
        {
            trace.Trace("GetNewListedPrice");
            EntityReference refPriceList = (EntityReference)enQuotation["bsd_pricelevel"];
            EntityReference refUnit = (EntityReference)enQuotation["bsd_unitnumber"];

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