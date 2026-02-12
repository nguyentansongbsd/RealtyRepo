// Decompiled with JetBrains decompiler
// Type: SaleDirectAction.SaleDirectAction
// Assembly: SaleDirectAction, Version=1.0.0.0, Culture=neutral, PublicKeyToken=4e71628980e853ee
// MVID: F9B79C4D-3E6B-49FD-A188-86807627C22E
// Assembly location: C:\Users\XUAN CHINH\Downloads\SaleDirectAction.dll

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IdentityModel.Metadata;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Web.UI.WebControls;

namespace SaleDirectAction
{
    public class SaleDirectAction : IPlugin
    {
        public IOrganizationService service;
        private IOrganizationServiceFactory factory;
        private StringBuilder strbuil = new StringBuilder();
        ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            EntityReference entityReference1 = (EntityReference)context.InputParameters["Target"];
            string str1 = context.InputParameters["Command"].ToString();
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                strbuil.AppendLine("11111111");
                //throw new InvalidPluginExecutionException(strbuil.ToString());
                Main(str1, entityReference1, context);
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }

        }
        public void Main(string str1, EntityReference entityReference1, IPluginExecutionContext context)
        {
            try
            {
                if (str1 == "Book")
                {
                    //factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    //service = factory.CreateOrganizationService(context.UserId);
                    Entity entity1 = RetrieveValidUnit(entityReference1.Id);
                    if (((OptionSetValue)entity1["statecode"]).Value == 1)
                        throw new InvalidPluginExecutionException("This unit is not public!");
                    if (((OptionSetValue)entity1["statuscode"]).Value == 100000002)
                        throw new InvalidPluginExecutionException("This unit was sold!");
                    if (!entity1.Contains("bsd_floor"))
                        throw new InvalidPluginExecutionException("Please select floor for this unit!");
                    if (!entity1.Contains("bsd_blocknumber"))
                        throw new InvalidPluginExecutionException("Please select block for this unit!");
                    if (!entity1.Contains("bsd_projectcode"))
                        throw new InvalidPluginExecutionException("Please select project for this unit!");
                    //if (!entity1.Contains("defaultuomid"))
                    //    throw new InvalidPluginExecutionException("Please select default unit for this unit!");

                    Entity entity2 = new Entity("bsd_opportunity");
                    entity2["bsd_name"] = (object)entity1["bsd_name"].ToString();
                    entity2["bsd_project"] = entity1["bsd_projectcode"];
                    entity2["bsd_unit"] = (object)entityReference1;
                    entity2["bsd_queueforproject"] = false;
                    entity2["bsd_pricelist"] = entity1.Contains("bsd_pricelevel") ? entity1["bsd_pricelevel"] : null;

                    Entity enPriceListItem = getPriceListItem(entity1);
                    if (enPriceListItem != null)
                    {
                        entity2["bsd_usableareasqm"] = enPriceListItem["bsd_netusablearea"];
                        entity2["bsd_builtupareasqm"] = enPriceListItem["bsd_builtuparea"];
                        entity2["bsd_usableunitprice"] = enPriceListItem["bsd_usableareaunitprice"];
                        entity2["bsd_builtupunitprice"] = enPriceListItem["bsd_builtupunitprice"];
                        entity2["bsd_price"] = enPriceListItem["bsd_price"];
                    }

                    DateTime dateNow = DateTime.Now;
                    entity2["bsd_bookingtime"] = RetrieveLocalTimeFromUTCTime(dateNow, service);

                    //EntityReference enfPhasesLaunch = entity1.Contains("bsd_pricelevel") ? PhasesLaunchPriceList((EntityReference)entity1["bsd_pricelevel"]) : null;
                    //if (((OptionSetValue)entity1["statuscode"]).Value == 100000000 && enfPhasesLaunch != null) // 100000000 = Available
                    //{
                    //    entity2["bsd_phaselaunch"] = entity1.Contains("bsd_pricelevel") ? PhasesLaunchPriceList((EntityReference)entity1["bsd_pricelevel"]) : null;
                    //    entity2["bsd_queuingfee"] = new Money(0);
                    //    entity2["statuscode"] = new OptionSetValue(100000006);

                    //    int shortTimeQueue = getShortTimeQueueByProject((EntityReference)entity1["bsd_projectcode"]);
                    //    DateTime bsd_queuingexpired = dateNow.AddHours(shortTimeQueue);
                    //    entity2["bsd_queuingexpired"] = RetrieveLocalTimeFromUTCTime(bsd_queuingexpired, service);
                    //}
                    //else
                    if ((((OptionSetValue)entity1["statuscode"]).Value == 100000000 )) //|| ((OptionSetValue)entity1["statuscode"]).Value == 100000004 && enfPhasesLaunch == null
                    {
                        EntityReference entityReference2 = (EntityReference)entity1["bsd_projectcode"];
                        Entity entity3 = service.Retrieve(entityReference2.LogicalName, entityReference2.Id, new ColumnSet(new string[] { "bsd_bookingfee" }));
                        if (entity3 == null)
                            throw new InvalidPluginExecutionException("Project named '" + entityReference2.Name + "' is not available!");
                        entity2["bsd_queuingfee"] = entity3.Contains("bsd_bookingfee") ? entity3["bsd_bookingfee"] : new Money(0);
                        entity2["statuscode"] = new OptionSetValue(100000006);

                        int longTimeQueue = getLongTimeQueueByProject((EntityReference)entity1["bsd_projectcode"]);
                        DateTime bsd_queuingexpired = dateNow.AddDays(longTimeQueue);
                        entity2["bsd_queuingexpired"] = RetrieveLocalTimeFromUTCTime(bsd_queuingexpired, service);
                        entity2["bsd_dateorder"] = RetrieveLocalTimeFromUTCTime(bsd_queuingexpired, service);
                    }

                    EntityReference pricelist_id = null;
                    //EntityReference pricelist_ref = null;

                    //if (pricelist_id != null)
                    //{
                    //    var rplCopy = getListByIDCopy(service, pricelist_id.Id);

                    //    if (rplCopy == null || rplCopy.Entities.Count == 0)
                    //    {
                    //    }
                    //    else
                    //    {
                    //        var copy = rplCopy[0];
                    //        pricelist_ref = new EntityReference(copy.LogicalName, copy.Id);
                    //    }
                    //    //  entity2["bsd_pricelistapply"] = pricelist_ref;
                    //}
                    //var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    //<fetch distinct=""true"">
                    //  <entity name=""bsd_productpricelevel"">
                    //    <filter>
                    //      <condition attribute=""bsd_product"" operator=""eq"" value=""{entity1.Id}"" />
                    //    </filter>
                    //    <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">
                    //      <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                    //        <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                    //          <attribute name=""bsd_name"" alias=""name"" />
                    //          <attribute name=""bsd_depositamount"" alias=""depositamount"" />
                    //          <attribute name=""bsd_minimumdeposit"" alias=""minimumdeposit"" />
                    //          <attribute name=""bsd_phaseslaunchid"" alias=""phaseid"" />
                    //          <filter>
                    //            <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                    //            <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                    //          </filter>
                    //        </link-entity>
                    //      </link-entity>
                    //    </link-entity>
                    //  </entity>
                    //</fetch>";
                    //EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    //if (rs.Entities.Count == 1)
                    //{
                    //    tracingService.Trace("vào if phase_" + rs.Entities.Count);

                    //    var aliased = (AliasedValue)rs.Entities[0]["phaseid"];
                    //    Guid phaseId = (Guid)aliased.Value;

                    //    entity2["bsd_phaseslaunchid"] = new EntityReference("bsd_phaseslaunch", phaseId);
                    //    var aliased_money = (AliasedValue)rs.Entities[0]["depositamount"];
                    //    Money moneyValue = (Money)aliased_money.Value;
                    //    entity2["bsd_depositfee"] = moneyValue;
                    //    var minimum = (AliasedValue)rs.Entities[0]["minimumdeposit"];
                    //    Money moneyminimum = (Money)minimum.Value;
                    //    entity2["bsd_minimumdeposit"] = moneyminimum;
                    //}
                    Guid guid = service.Create(entity2);
                    //if (((OptionSetValue)entity1["statuscode"]).Value == 100000000 && enfPhasesLaunch != null)
                    //    updateUnitStatus(entityReference1, 100000004);

                    //Entity entity4 = new Entity("opportunityproduct");
                    //entity4["opportunityid"] = entity4["bsd_booking"] = new EntityReference("opportunity", guid);
                    //entity4["uomid"] = entity1["defaultuomid"];
                    //entity4["bsd_floor"] = entity1["bsd_floor"];
                    //entity4["bsd_block"] = entity1["bsd_blocknumber"];
                    //entity4["bsd_project"] = entity1["bsd_projectcode"];
                    //entity4["bsd_productid"] = entity4["bsd_units"] = entityReference1;
                    //entity4["isproductoverridden"] = false;
                    //entity4["ispriceoverridden"] = false;
                    //StringBuilder st = new StringBuilder();
                    //decimal amount = 0;
                    //if (pricelist_ref != null)
                    //{

                    //    var fetchXml = $@"
                    //                <fetch>
                    //                  <entity name='productpricelevel'>
                    //                    <attribute name='amount' />
                    //                    <filter>
                    //                      <condition attribute='pricelevelid' operator='eq' value='{pricelist_ref.Id}'/>
                    //                      <condition attribute='productid' operator='eq' value='{entityReference1.Id}'/>
                    //                    </filter>
                    //                  </entity>
                    //                </fetch>";
                    //    EntityCollection list = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    //    if (list.Entities.Count > 0)
                    //    {
                    //        amount = list.Entities[0].Contains("amount") ? ((Money)list.Entities[0]["amount"]).Value : 0;
                    //        if (amount > 0)
                    //        {
                    //            entity4["isproductoverridden"] = true;
                    //            entity4["ispriceoverridden"] = true;
                    //            entity4["priceperunit"] = new Money(amount);
                    //            entity4["extendedamount"] = new Money(amount);

                    //            entity4["bsd_pricelist"] = pricelist_ref;
                    //        }

                    //    }



                    //}
                    //else
                    //{
                    //    if (entity1.Contains("bsd_listprice"))
                    //        entity4["priceperunit"] = entity1["bsd_listprice"];

                    //    if (entity2.Contains("pricelevelid"))
                    //        entity4["bsd_pricelist"] = entity2["pricelevelid"];

                    //}


                    //entity4["quantity"] = (object)Decimal.One;

                    //if (entity1.Contains("bsd_phaseslaunchid"))
                    //{
                    //    entity4["bsd_status"] = true;
                    //    entity4["bsd_phaseslaunch"] = entity1["bsd_phaseslaunchid"];
                    //}
                    ////throw new InvalidPluginExecutionException("nghiax tesst ");
                    //Guid idopppro = service.Create(entity4);



                    string str2 = "tmp={type:'Success',content:'" + guid.ToString() + "'}";
                    context.OutputParameters["Result"] = str2;
                }
                else if (str1 == "Quotation")
                {
                    //factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    //service = factory.CreateOrganizationService(context.UserId);
                    tracingService.Trace("2");
                    Entity enUnit = RetrieveValidUnit(entityReference1.Id);
                    if (enUnit == null)
                        throw new InvalidPluginExecutionException("Unit is not avaliable please check detail of unit!");
                    if (((OptionSetValue)enUnit["statecode"]).Value == 1)
                        throw new InvalidPluginExecutionException("This unit is not public!");
                    if (((OptionSetValue)enUnit["statuscode"]).Value != 100000000 && ((OptionSetValue)enUnit["statuscode"]).Value != 100000001 && ((OptionSetValue)enUnit["statuscode"]).Value != 100000004)
                        throw new InvalidPluginExecutionException("Unit must be available or queueing!");
                    if (((OptionSetValue)enUnit["statuscode"]).Value == 100000002)
                        throw new InvalidPluginExecutionException("Unit is sold!");
                    //if (!enUnit.Contains("bsd_phaseslaunchid"))
                    //    throw new InvalidPluginExecutionException("Unit is not launched!");
                    if (!enUnit.Contains("bsd_floor"))
                        throw new InvalidPluginExecutionException("Please select floor for this unit!");
                    if (!enUnit.Contains("bsd_blocknumber"))
                        throw new InvalidPluginExecutionException("Please select block for this unit!");
                    if (!enUnit.Contains("bsd_projectcode"))
                        throw new InvalidPluginExecutionException("Please select project for this unit!");
                    //if (!enUnit.Contains("defaultuomid"))
                    //    throw new InvalidPluginExecutionException("Please select default unit for this unit!");

                    tracingService.Trace("3");
                    Entity entity2 = new Entity("bsd_quote");
                    tracingService.Trace("3.1");
                    entity2["bsd_name"] = enUnit["bsd_name"];
                    tracingService.Trace("3.3");
                    entity2["bsd_projectid"] = enUnit["bsd_projectcode"];
                    entity2["transactioncurrencyid"] = enUnit["transactioncurrencyid"];

                    tracingService.Trace("3.4");
                    //entity2["bsd_phaseslaunchid"] = enUnit["bsd_phaseslaunchid"];
                    entity2["bsd_unitno"] = (object)entityReference1;
                    entity2["statuscode"] = new OptionSetValue(100000007);
                    strbuil.AppendLine("22222");
                    entity2["bsd_netusablearea"] = enUnit.Contains("bsd_netsaleablearea") ? enUnit["bsd_netsaleablearea"] : Decimal.Zero;
                    entity2["bsd_constructionarea"] = enUnit.Contains("bsd_constructionarea") ? enUnit["bsd_constructionarea"] : Decimal.Zero;
                    int numberofmonthspaidmf = -1;
                    strbuil.AppendLine("333333");
                    Entity entity3 = service.Retrieve(((EntityReference)enUnit["bsd_projectcode"]).LogicalName, ((EntityReference)enUnit["bsd_projectcode"]).Id, new ColumnSet(true));
                    if (enUnit.Contains("bsd_numberofmonthspaidmf"))
                    {
                        numberofmonthspaidmf = (int)enUnit["bsd_numberofmonthspaidmf"];
                        entity2["bsd_numberofmonthspaidmf"] = enUnit["bsd_numberofmonthspaidmf"];
                    }
                    else if (entity3.Contains("bsd_numberofmonthspaidmf"))
                    {
                        numberofmonthspaidmf = (int)entity3["bsd_numberofmonthspaidmf"];
                        entity2["bsd_numberofmonthspaidmf"] = entity3["bsd_numberofmonthspaidmf"];
                    }

                    entity2["bsd_totalamountpaid"] = new Money(0);
                    DateTime utcNow = DateTime.UtcNow;
                    DateTime localNow = RetrieveLocalTimeFromUTCTime(utcNow, service);
                    entity2["bsd_quotationdate"] = localNow;
                    strbuil.AppendLine("444444");
                    tracingService.Trace("4");
                    Decimal managementamount = entity3.Contains("bsd_managementamount") ? ((Money)entity3["bsd_managementamount"]).Value : Decimal.Zero;
                    Decimal netsaleablearea = enUnit.Contains("bsd_netsaleablearea") ? (Decimal)enUnit["bsd_netsaleablearea"] : Decimal.Zero;
                    Decimal actualarea = enUnit.Contains("bsd_actualarea") ? (Decimal)enUnit["bsd_actualarea"] : Decimal.Zero;
                    if (enUnit.Contains("bsd_managementamountmonth"))
                        managementamount = ((Money)enUnit["bsd_managementamountmonth"]).Value;
                    if (numberofmonthspaidmf > -1)
                    {
                        Decimal num5 = netsaleablearea;
                        Decimal num6 = new Decimal(1.1);
                        Decimal num7 = (Decimal)numberofmonthspaidmf * num5 * managementamount;
                        entity2["bsd_managementfee"] = (object)new Money(num7);
                    }
                    strbuil.AppendLine("5555555555");
                    tracingService.Trace("5");
                    if (enUnit.Contains("bsd_taxcode"))
                    {
                        entity2["bsd_taxcode"] = enUnit["bsd_taxcode"];

                    }
                    if (enUnit.Contains("bsd_maintenancefeespercent"))
                    {
                        entity2["bsd_maintenancefeespercent"] = enUnit["bsd_maintenancefeespercent"];

                    }
                    if (enUnit.Contains("bsd_maintenancefees"))
                    {
                        entity2["bsd_maintenancefees"] = enUnit["bsd_maintenancefees"];

                    }
                    tracingService.Trace("6");
                    #region Update pricelist mới nhất
                    EntityReference pricelist_ref = null;
                    tracingService.Trace("6....");
                    if (entity2.Contains("bsd_pricelistphaselaunch") && entity2["bsd_pricelistphaselaunch"] != null)
                    {
                        tracingService.Trace("6.033333");
                        pricelist_ref = (EntityReference)entity2["bsd_pricelistphaselaunch"];
                    }

                    tracingService.Trace("6.0");
                    strbuil.AppendLine("aaaaa");
                    if (pricelist_ref != null)
                    {
                        tracingService.Trace("6.01");
                        strbuil.AppendLine("bbbbbbb");
                        var rplCopy = getListByIDCopy(service, pricelist_ref.Id);
                        strbuil.AppendLine("cccccc");
                        if (rplCopy == null || rplCopy.Entities.Count == 0)
                        {
                            strbuil.AppendLine("ddddd");
                        }
                        else
                        {
                            strbuil.AppendLine("eeeee");
                            var copy = rplCopy[0];
                            pricelist_ref = new EntityReference(copy.LogicalName, copy.Id);
                        }
                        strbuil.AppendLine("fffffffff");
                        entity2["pricelevelid"] = pricelist_ref;
                    }
                    #endregion
                    strbuil.AppendLine("7777777");
                    tracingService.Trace("6.1");
                    tracingService.Trace("7");
                    if (context.InputParameters.Contains("Parameters") && context.InputParameters["Parameters"] != null)
                    {
                        strbuil.AppendLine("88888888");
                        tracingService.Trace("8");
                        DataContractJsonSerializer contractJsonSerializer = new DataContractJsonSerializer(typeof(InputParameter[]));
                        MemoryStream ser = new MemoryStream(Encoding.UTF8.GetBytes((string)context.InputParameters["Parameters"]));
                        InputParameter[] inputParameter1 = (InputParameter[])contractJsonSerializer.ReadObject(ser);
                        foreach (InputParameter inputParameter in inputParameter1)
                        {
                            if (inputParameter.action == str1)
                            {
                                strbuil.AppendLine("9999999999");
                                tracingService.Trace("9");
                                Entity entity4 = service.Retrieve(inputParameter.name, Guid.Parse(inputParameter.value), new ColumnSet(new string[5]
                                {
                  "bsd_nameofstaffagent",
                  "customerid",
                  "bsd_queuingfee",
                  "bsd_salesagentcompany",
                  "bsd_referral"
                                }));
                                EntityReference entityReference2 = entity4.Contains("customerid") ? (EntityReference)entity4["customerid"] : (EntityReference)null;
                                if (entityReference2 != null)
                                {
                                    tracingService.Trace("10");
                                    entity2["customerid"] = (object)entityReference2;
                                    EntityReference enfBA = getBankAccount(entityReference2.Id);
                                    tracingService.Trace("11");
                                    if (enfBA != null)
                                        entity2["bsd_bankaccount"] = enfBA;
                                }
                                if (entity4.Contains("bsd_queuingfee"))
                                    entity2["bsd_bookingfee"] = entity4["bsd_queuingfee"];
                                if (entity4.Contains("bsd_nameofstaffagent"))
                                    entity2["bsd_nameofstaffagent"] = entity4["bsd_nameofstaffagent"];
                                if (entity4.Contains("bsd_salesagentcompany"))
                                    entity2["bsd_salessgentcompany"] = entity4["bsd_salesagentcompany"];
                                if (entity4.Contains("bsd_referral"))
                                    entity2["bsd_referral"] = entity4["bsd_referral"];
                            }
                        }

                    }
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <attribute name=""bsd_name"" alias=""name"" />
                              <attribute name=""bsd_depositamount"" alias=""depositamount"" />
                              <attribute name=""bsd_minimumdeposit"" alias=""minimumdeposit"" />
                              <attribute name=""bsd_phaseslaunchid"" alias=""phaseid"" />
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (rs.Entities.Count == 1)
                    {
                        tracingService.Trace("vào if phase_" + rs.Entities.Count);

                        var aliased = (AliasedValue)rs.Entities[0]["phaseid"];
                        Guid phaseId = (Guid)aliased.Value;

                        entity2["bsd_phaseslaunchid"] = new EntityReference("bsd_phaseslaunch", phaseId);
                        var aliased_money = (AliasedValue)rs.Entities[0]["depositamount"];
                        Money moneyValue = (Money)aliased_money.Value;
                        entity2["bsd_depositfee"] = moneyValue;
                        var minimum = (AliasedValue)rs.Entities[0]["minimumdeposit"];
                        Money moneyminimum = (Money)minimum.Value;
                        entity2["bsd_minimumdeposit"] = moneyminimum;
                    }
                    //var fetchXml_pricelist = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    //<fetch distinct=""true"">
                    //  <entity name=""bsd_productpricelevel"">
                    //    <attribute name=""bsd_price"" alias=""prilist_price"" />
                    //    <filter>
                    //      <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                    //    </filter>
                    //    <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                    //      <attribute name=""bsd_name"" alias=""price_name"" />
                    //      <attribute name=""bsd_pricelevelid"" alias=""price_id"" />
                    //      <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                    //        <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                    //          <filter>
                    //            <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                    //            <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                    //          </filter>
                    //        </link-entity>
                    //      </link-entity>
                    //    </link-entity>
                    //  </entity>
                    //</fetch>";
                    //EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml_pricelist));
                    //if (rs_price.Entities.Count == 1)
                    //{
                    //    tracingService.Trace("vào if price_" + rs_price.Entities.Count);

                    //    var aliased_price = (AliasedValue)rs_price.Entities[0]["price_id"];
                    //    Guid price_id = (Guid)aliased_price.Value;
                    //    entity2["bsd_pricelevel"] = new EntityReference("bsd_pricelevel", price_id);
                    //    if (rs_price.Entities[0].Contains("prilist_price"))
                    //    {
                    //        tracingService.Trace("Có prilist_price");
                    //        var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
                    //        Money moneyValue = (Money)aliased_money.Value;

                    //        entity2["bsd_detailamount"] = moneyValue;
                    //        if (enUnit.Contains("bsd_taxcode"))
                    //        {
                    //            tracingService.Trace("Có bsd_taxcode");
                    //            Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
                    //            decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                    //            decimal taxRate = taxCodeValue / 100.0m;
                    //            decimal detailAmount = moneyValue.Value;
                    //            decimal vatAmount = detailAmount * taxRate;
                    //            entity2["bsd_vat"] = new Money(vatAmount);
                    //        }
                    //    }
                    //}

                    entity2["bsd_pricelevel"] = enUnit.Contains("bsd_pricelevel") ? (EntityReference)enUnit["bsd_pricelevel"] : null;
                    if (enUnit.Contains("bsd_pricelevel") && enUnit["bsd_pricelevel"] != null)
                    {
                        EntityReference priceLevelRef = (EntityReference)enUnit["bsd_pricelevel"];

                        var fetchXml123 = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
                              <entity name=""bsd_productpricelevel"">
                                <attribute name=""bsd_price"" alias=""p_price"" />
                                <filter type=""and"">
                                  <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                                  <condition attribute=""statuscode"" operator=""ne"" value=""{2}"" />
                                </filter>
                                <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                                  <filter type=""and"">
                                    <condition attribute=""bsd_pricelevelid"" operator=""eq"" value=""{priceLevelRef.Id}"" />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                        EntityCollection rs_price123 = service.RetrieveMultiple(new FetchExpression(fetchXml123));
                        if (rs_price123.Entities.Count > 0)
                        {
                            tracingService.Trace("vào if price_" + rs_price123.Entities.Count);
                            if (rs_price123.Entities[0].Contains("p_price"))
                            {
                                Entity result = rs_price123.Entities[0];
                                tracingService.Trace("Có p_price");
                                AliasedValue aliasedPrice = (AliasedValue)result["p_price"];


                                Money moneyValue = (Money)aliasedPrice.Value;
                                entity2["bsd_detailamount"] = moneyValue;
                                if (enUnit.Contains("bsd_taxcode"))
                                {
                                    tracingService.Trace("Có bsd_taxcode");
                                    Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
                                    decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                                    decimal taxRate = taxCodeValue / 100.0m;
                                    decimal detailAmount = moneyValue.Value;
                                    decimal vatAmount = detailAmount * taxRate;
                                    entity2["bsd_vat"] = new Money(vatAmount);
                                }
                            }
                        }
                    }

                    tracingService.Trace("12");
                    strbuil.AppendLine("aaaaa");

                    int nextNumber = 1;
                    string fetchMaxCode = $@"
                    <fetch top='1'>
                      <entity name='bsd_quote'>
                        <attribute name='bsd_reservationno' />
                        <order attribute='bsd_reservationno' descending='true' />
                      </entity>
                    </fetch>";
                    EntityCollection lastRecords = service.RetrieveMultiple(new FetchExpression(fetchMaxCode));

                    if (lastRecords.Entities.Count > 0)
                    {
                        // Lấy chuỗi RSC-00000001
                        string lastCode = lastRecords.Entities[0].GetAttributeValue<string>("bsd_reservationno");

                        if (!string.IsNullOrEmpty(lastCode))
                        {
                            // Cắt bỏ phần chữ "RSC-", chỉ lấy phần số "00000001"
                            string numericPart = lastCode.Replace("RSV-", "");
                            if (int.TryParse(numericPart, out int lastNumber))
                            {
                                nextNumber = lastNumber + 1;
                            }
                        }
                    }
                    // Gán mã mới vào entity: Ví dụ RSC-00000002
                    entity2["bsd_reservationno"] = "RSV-" + nextNumber.ToString("D7");
                    Guid guid = service.Create(entity2);
                    create_update_DataProjection(entityReference1.Id, entity2, guid);
                    tracingService.Trace("16");
                    strbuil.AppendLine("bbbbbbbbbbbb");
                    //throw new InvalidPluginExecutionException(strbuil.ToString());
                    context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + guid.ToString() + "'}";
                    tracingService.Trace("Done quotation");
                }
                else if (str1 == "Reservation")
                {

                    Entity enUnit = RetrieveValidUnit(entityReference1.Id);
                    Entity updateUnit = new Entity("bsd_product", entityReference1.Id);
                    Entity entity2 = new Entity("bsd_quote");
                    var fetchXml1 = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <attribute name=""bsd_price"" alias=""prilist_price"" />
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <attribute name=""bsd_name"" alias=""name"" />
                              <attribute name=""bsd_depositamount"" alias=""depositamount"" />
                              <attribute name=""bsd_minimumdeposit"" alias=""minimumdeposit"" />
                              <attribute name=""bsd_phaseslaunchid"" alias=""phaseid"" />
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml1));
                    tracingService.Trace("phase_" + rs.Entities.Count);
                    if (rs.Entities.Count == 1)
                    {
                        tracingService.Trace("vào if phase_" + rs.Entities.Count);

                        var aliased = (AliasedValue)rs.Entities[0]["phaseid"];
                        Guid phaseId = (Guid)aliased.Value;

                        entity2["bsd_phaseslaunchid"] = new EntityReference("bsd_phaseslaunch", phaseId);
                        var aliased_money = (AliasedValue)rs.Entities[0]["depositamount"];
                        Money moneyValue = (Money)aliased_money.Value;
                        entity2["bsd_depositfee"] = moneyValue;
                        var minimum = (AliasedValue)rs.Entities[0]["minimumdeposit"];
                        Money moneyminimum = (Money)minimum.Value;
                        entity2["bsd_minimumdeposit"] = moneyminimum;
                    }
                    //var fetchXml_pricelist = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    //<fetch distinct=""true"">
                    //  <entity name=""bsd_productpricelevel"">
                    //    <attribute name=""bsd_price"" alias=""prilist_price"" />
                    //    <filter>
                    //      <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                    //    </filter>
                    //    <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                    //      <attribute name=""bsd_name"" alias=""price_name"" />
                    //      <attribute name=""bsd_pricelevelid"" alias=""price_id"" />
                    //      <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                    //        <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                    //          <filter>
                    //            <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                    //            <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                    //          </filter>
                    //        </link-entity>
                    //      </link-entity>
                    //    </link-entity>
                    //  </entity>
                    //</fetch>";
                    //EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml_pricelist));
                    //if (rs_price.Entities.Count == 1)
                    //{
                    //    tracingService.Trace("vào if price_" + rs_price.Entities.Count);

                    //    var aliased_price = (AliasedValue)rs_price.Entities[0]["price_id"];
                    //    Guid price_id = (Guid)aliased_price.Value;
                    //    entity2["bsd_pricelevel"] = new EntityReference("bsd_pricelevel", price_id);
                    //    if (rs_price.Entities[0].Contains("prilist_price"))
                    //    {
                    //        var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
                    //        Money moneyValue = (Money)aliased_money.Value;

                    //        entity2["bsd_detailamount"] = moneyValue;
                    //        if (enUnit.Contains("bsd_taxcode"))
                    //        {
                    //            Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
                    //            decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                    //            decimal taxRate = taxCodeValue / 100.0m;
                    //            decimal detailAmount = moneyValue.Value;
                    //            decimal vatAmount = detailAmount * taxRate;
                    //            entity2["bsd_vat"] = new Money(vatAmount);
                    //        }
                    //    }
                    //}
                    entity2["bsd_pricelevel"] = enUnit.Contains("bsd_pricelevel") ? (EntityReference)enUnit["bsd_pricelevel"] : null;
                    if (enUnit.Contains("bsd_pricelevel") && enUnit["bsd_pricelevel"] != null)
                    {
                        EntityReference priceLevelRef = (EntityReference)enUnit["bsd_pricelevel"];

                        var fetchXml123 = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
                              <entity name=""bsd_productpricelevel"">
                                <attribute name=""bsd_price"" alias=""p_price"" />
                                <filter type=""and"">
                                  <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                                  <condition attribute=""statuscode"" operator=""ne"" value=""{2}"" />
                                </filter>
                                <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                                  <filter type=""and"">
                                    <condition attribute=""bsd_pricelevelid"" operator=""eq"" value=""{priceLevelRef.Id}"" />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                        EntityCollection rs_price123 = service.RetrieveMultiple(new FetchExpression(fetchXml123));
                        if (rs_price123.Entities.Count > 0)
                        {
                            tracingService.Trace("vào if price_" + rs_price123.Entities.Count);
                            if (rs_price123.Entities[0].Contains("p_price"))
                            {
                                Entity result = rs_price123.Entities[0];
                                tracingService.Trace("Có p_price");
                                AliasedValue aliasedPrice = (AliasedValue)result["p_price"];


                                Money moneyValue = (Money)aliasedPrice.Value;
                                entity2["bsd_detailamount"] = moneyValue;
                                if (enUnit.Contains("bsd_taxcode"))
                                {
                                    tracingService.Trace("Có bsd_taxcode");
                                    Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
                                    decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                                    decimal taxRate = taxCodeValue / 100.0m;
                                    decimal detailAmount = moneyValue.Value;
                                    decimal vatAmount = detailAmount * taxRate;
                                    entity2["bsd_vat"] = new Money(vatAmount);
                                }
                            }
                        }
                    }

                    //
                    updateUnit["statuscode"] = new OptionSetValue(100000003);
                    service.Update(updateUnit);
                    entity2["bsd_name"] = enUnit["bsd_name"];
                    entity2["bsd_projectid"] = enUnit["bsd_projectcode"];
                    entity2["transactioncurrencyid"] = enUnit["transactioncurrencyid"];
                    entity2["bsd_unitno"] = (object)entityReference1;
                    entity2["statuscode"] = new OptionSetValue(100000000);
                    if (enUnit.Contains("bsd_taxcode"))
                    {
                        entity2["bsd_taxcode"] = enUnit["bsd_taxcode"];
                    }
                    if (enUnit.Contains("bsd_maintenancefeespercent"))
                    {
                        entity2["bsd_maintenancefeespercent"] = enUnit["bsd_maintenancefeespercent"];

                    }
                    if (enUnit.Contains("bsd_maintenancefees"))
                    {
                        entity2["bsd_maintenancefees"] = enUnit["bsd_maintenancefees"];

                    }
                    entity2["bsd_reservationtime"] = DateTime.Today;
                    entity2["bsd_netusablearea"] = enUnit.Contains("bsd_netsaleablearea") ? enUnit["bsd_netsaleablearea"] : Decimal.Zero;
                    entity2["bsd_constructionarea"] = enUnit.Contains("bsd_constructionarea") ? enUnit["bsd_constructionarea"] : Decimal.Zero;
                    Entity entity3 = service.Retrieve(((EntityReference)enUnit["bsd_projectcode"]).LogicalName, ((EntityReference)enUnit["bsd_projectcode"]).Id, new ColumnSet(true));
                    entity2["bsd_totalamountpaid"] = new Money(0);
                    //DateTime utcNow = DateTime.UtcNow;
                    //DateTime localNow = RetrieveLocalTimeFromUTCTime(utcNow, service);
                    //entity2["bsd_quotationdate"] = localNow;
                    int nextNumber = 1;
                    string fetchMaxCode = $@"
                    <fetch top='1'>
                      <entity name='bsd_quote'>
                        <attribute name='bsd_reservationno' />
                        <order attribute='bsd_reservationno' descending='true' />
                      </entity>
                    </fetch>";
                    EntityCollection lastRecords = service.RetrieveMultiple(new FetchExpression(fetchMaxCode));

                    if (lastRecords.Entities.Count > 0)
                    {
                        // Lấy chuỗi RSC-00000001
                        string lastCode = lastRecords.Entities[0].GetAttributeValue<string>("bsd_reservationno");

                        if (!string.IsNullOrEmpty(lastCode))
                        {
                            // Cắt bỏ phần chữ "RSC-", chỉ lấy phần số "00000001"
                            string numericPart = lastCode.Replace("RSV-", "");
                            if (int.TryParse(numericPart, out int lastNumber))
                            {
                                nextNumber = lastNumber + 1;
                            }
                        }
                    }
                    // Gán mã mới vào entity: Ví dụ RSC-00000002
                    entity2["bsd_reservationno"] = "RSV-" + nextNumber.ToString("D7");
                    Guid guid = service.Create(entity2);
                    create_update_DataProjection(entityReference1.Id, entity2, guid);
                    context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + guid.ToString() + "'}";
                }
                else if (str1 == "RAContract")
                {
                    Entity enUnit = RetrieveValidUnit(entityReference1.Id);
                    Entity updateUnit = new Entity("bsd_product", entityReference1.Id);
                    Entity enReContract = new Entity("bsd_reservationcontract");
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <attribute name=""bsd_price"" alias=""prilist_price"" />
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <attribute name=""bsd_name"" alias=""name"" />
                              <attribute name=""bsd_depositamount"" alias=""depositamount"" />
                              <attribute name=""bsd_phaseslaunchid"" alias=""phaseid"" />
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (rs.Entities.Count == 1)
                    {
                        tracingService.Trace("vào if phase_" + rs.Entities.Count);

                        var aliased = (AliasedValue)rs.Entities[0]["phaseid"];
                        Guid phaseId = (Guid)aliased.Value;

                        enReContract["bsd_phaseslaunchid"] = new EntityReference("bsd_phaseslaunch", phaseId);
                        var aliased_money = (AliasedValue)rs.Entities[0]["depositamount"];
                        Money moneyValue = (Money)aliased_money.Value;
                        //enReContract["bsd_depositfee"] = moneyValue;
                    }
                    var fetchXml_pricelist = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_productpricelevel"">
                        <attribute name=""bsd_price"" alias=""prilist_price"" />
                        <filter>
                          <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                        </filter>
                        <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"" alias=""price"">
                          <attribute name=""bsd_name"" alias=""price_name"" />
                          <attribute name=""bsd_pricelevelid"" alias=""price_id"" />
                          <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                            <link-entity name=""bsd_phaseslaunch"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""phase"" intersect=""true"">
                              <filter>
                                <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                                <condition attribute=""bsd_stopselling"" operator=""eq"" value=""{0}"" />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml_pricelist));
                    if (rs_price.Entities.Count == 1)
                    {
                        tracingService.Trace("vào if price_" + rs_price.Entities.Count);

                        var aliased_price = (AliasedValue)rs_price.Entities[0]["price_id"];
                        Guid price_id = (Guid)aliased_price.Value;
                        enReContract["bsd_pricelevel"] = new EntityReference("bsd_pricelevel", price_id);
                        if (rs_price.Entities[0].Contains("prilist_price"))
                        {
                            var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
                            Money moneyValue = (Money)aliased_money.Value;

                            enReContract["bsd_detailamount"] = moneyValue;
                            if (enUnit.Contains("bsd_taxcode"))
                            {
                                Entity entity_taxcode = service.Retrieve(((EntityReference)enUnit["bsd_taxcode"]).LogicalName, ((EntityReference)enUnit["bsd_taxcode"]).Id, new ColumnSet(true));
                                decimal taxCodeValue = entity_taxcode.Contains("bsd_value") ? (decimal)entity_taxcode["bsd_value"] : 0;
                                decimal taxRate = taxCodeValue / 100.0m;
                                decimal detailAmount = moneyValue.Value;
                                decimal vatAmount = detailAmount * taxRate;
                                //entity2["bsd_vat"] = new Money(vatAmount);
                            }
                        }
                    }
                    updateUnit["statuscode"] = new OptionSetValue(100000006);//Reserver
                    service.Update(updateUnit);
                    enReContract["bsd_name"] = enUnit["bsd_name"];
                    enReContract["bsd_projectid"] = enUnit["bsd_projectcode"];
                    //entity2["transactioncurrencyid"] = enUnit["transactioncurrencyid"];
                    enReContract["bsd_unitno"] = (object)entityReference1;

                    enReContract["statuscode"] = new OptionSetValue(1);
                    enReContract["bsd_approvalstatus"] = new OptionSetValue(100000000);
                    enReContract["bsd_landvaluededuction"] = new Money(0);
                    if (enUnit.Contains("bsd_taxcode"))
                    {
                        enReContract["bsd_taxcode"] = enUnit["bsd_taxcode"];
                    }
                    if (enUnit.Contains("bsd_numberofmonthspaidmf"))
                    {
                        enReContract["bsd_numberofmonthspaidmf"] = enUnit["bsd_numberofmonthspaidmf"];
                    }
                    if (enUnit.Contains("bsd_managementfee"))
                    {
                        enReContract["bsd_managementfee"] = enUnit["bsd_maintenancefeespercent"];

                    }
                    if (enUnit.Contains("bsd_unittype"))
                    {
                        enReContract["bsd_unittype"] = enUnit["bsd_unittype"];

                    }
                    //entity2["bsd_reservationtime"] = DateTime.Today;
                    enReContract["bsd_netusablearea"] = enUnit.Contains("bsd_netsaleablearea") ? enUnit["bsd_netsaleablearea"] : Decimal.Zero;
                    enReContract["bsd_constructionarea"] = enUnit.Contains("bsd_constructionarea") ? enUnit["bsd_constructionarea"] : Decimal.Zero;
                    //Entity entity3 = service.Retrieve(((EntityReference)enUnit["bsd_projectcode"]).LogicalName, ((EntityReference)enUnit["bsd_projectcode"]).Id, new ColumnSet(true));
                    int nextNumber = 1;
                    string fetchMaxCode = $@"
                    <fetch top='1'>
                      <entity name='bsd_reservationcontract'>
                        <attribute name='bsd_reservationnumber' />
                        <order attribute='bsd_reservationnumber' descending='true' />
                      </entity>
                    </fetch>";

                    EntityCollection lastRecords = service.RetrieveMultiple(new FetchExpression(fetchMaxCode));

                    if (lastRecords.Entities.Count > 0)
                    {
                        // Lấy chuỗi RSC-00000001
                        string lastCode = lastRecords.Entities[0].GetAttributeValue<string>("bsd_reservationnumber");

                        if (!string.IsNullOrEmpty(lastCode))
                        {
                            // Cắt bỏ phần chữ "RSC-", chỉ lấy phần số "00000001"
                            string numericPart = lastCode.Replace("RSC-", "");
                            if (int.TryParse(numericPart, out int lastNumber))
                            {
                                nextNumber = lastNumber + 1;
                            }
                        }
                    }
                    // Gán mã mới vào entity: Ví dụ RSC-00000002
                    enReContract["bsd_racontractsigndate"] = DateTime.Today;
                    enReContract["bsd_reservationnumber"] = "RSC-" + nextNumber.ToString("D8");
                    Guid guid = service.Create(enReContract);
                    create_update_DataProjection(entityReference1.Id, enReContract, guid);
                    context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + guid.ToString() + "'}";
                }
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        private void create_update_DataProjection(Guid idUnit, Entity enEntity, Guid id)
        {
            // get DataProjection theo unit
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_dataprojection"">
                <filter>
                  <condition attribute=""bsd_productid"" operator=""eq"" value=""{idUnit}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection en = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (en.Entities.Count > 0)
            {
                Entity enDataprojection = en.Entities[0];
                Entity enUp = new Entity(enDataprojection.LogicalName, enDataprojection.Id);
                if (enEntity.LogicalName == "bsd_quote") enUp["bsd_depositid"] = new EntityReference("bsd_quote", id);
                if (enEntity.LogicalName == "bsd_reservationcontract") enUp["bsd_raid"] = new EntityReference("bsd_reservationcontract", id);
                if (enEntity.Contains("bsd_customerid")) enUp["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_projectid")) enUp["bsd_project"] = enEntity["bsd_projectid"];
                if (enEntity.Contains("bsd_opportunityid")) enUp["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_queue")) enUp["bsd_bookingid"] = enEntity["bsd_queue"];
                if (enEntity.Contains("bsd_phaseslaunchid")) enUp["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunchid"];
                if (enEntity.Contains("bsd_quoteid")) enUp["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Update(enUp);
            }
            else
            {
                Entity enCre = new Entity("bsd_dataprojection");
                if (enEntity.LogicalName == "bsd_quote") enCre["bsd_depositid"] = new EntityReference("bsd_quote", id);
                if (enEntity.LogicalName == "bsd_reservationcontract") enCre["bsd_raid"] = new EntityReference("bsd_reservationcontract", id);
                if (enEntity.Contains("bsd_customerid")) enCre["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_projectid")) enCre["bsd_project"] = enEntity["bsd_projectid"];
                if (enEntity.Contains("bsd_opportunityid")) enCre["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_queue")) enCre["bsd_bookingid"] = enEntity["bsd_queue"];
                if (enEntity.Contains("bsd_phaseslaunchid")) enCre["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunchid"];
                if (enEntity.Contains("bsd_unitno")) enCre["bsd_productid"] = enEntity["bsd_unitno"];
                if (enEntity.Contains("bsd_quoteid")) enCre["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Create(enCre);
            }
        }

        private int getLongTimeQueueByProject(EntityReference enfProject)
        {
            Entity project = service.Retrieve(enfProject.LogicalName, enfProject.Id, new ColumnSet(new string[] {
                "bsd_longqueuingtime"
            }));
            if (project == null)
                return 0;
            else
                return project.Contains("bsd_longqueuingtime") ? (int)project["bsd_longqueuingtime"] : 0;
        }
        private int getShortTimeQueueByProject(EntityReference enfProject)
        {
            Entity project = service.Retrieve(enfProject.LogicalName, enfProject.Id, new ColumnSet(new string[] {
                "bsd_shortqueingtime"
            }));
            if (project == null)
                return 0;
            else
                return project.Contains("bsd_shortqueingtime") ? (int)project["bsd_shortqueingtime"] : 0;
        }
        private void updateUnitStatus(EntityReference unitRef, int statuscode)
        {
            Entity updateUnit = new Entity("bsd_product", unitRef.Id);
            updateUnit["statuscode"] = new OptionSetValue(statuscode);
            service.Update(updateUnit);
        }
        private Entity getPriceListItem(Entity enUnit)
        {
            if (!enUnit.Contains("bsd_pricelevel")) return null;
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_productpricelevel"">
                <attribute name=""bsd_builtuparea"" />
                <attribute name=""bsd_builtupunitprice"" />
                <attribute name=""bsd_netusablearea"" />
                <attribute name=""bsd_price"" />
                <attribute name=""bsd_usableareaunitprice"" />
                <filter>
                  <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                  <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{((EntityReference)enUnit["bsd_pricelevel"]).Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection entcs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (entcs.Entities.Count == 0)
                return null;
            else
                return entcs.Entities[0];
        }
        private EntityCollection findTaxCode()
        {
            string str = string.Format("<fetch version='1.0' output-format='xml-platform' count='1' mapping='logical' distinct='false'>\r\n                      <entity name='bsd_taxcode'>\r\n                        <attribute name='bsd_taxcodeid' />\r\n                        <filter type='and'>\r\n                          <condition attribute='bsd_default' operator='eq' value='1' />\r\n                        </filter>\r\n                      </entity>\r\n                    </fetch>");
            return service.RetrieveMultiple((QueryBase)new FetchExpression(str));
        }

        private Entity RetrieveValidUnit(Guid unitId)
        {
            QueryExpression q = new QueryExpression();
            q.EntityName = "bsd_product";
            q.ColumnSet = new ColumnSet(true);
            q.Criteria = new FilterExpression(LogicalOperator.And);
            q.Criteria.AddCondition(new ConditionExpression("bsd_productid", ConditionOperator.Equal, unitId));
            LinkEntity link_floor_unit = new LinkEntity("bsd_product", "bsd_floor", "bsd_floor", "bsd_floorid", JoinOperator.Inner);
            link_floor_unit.EntityAlias = "fl";
            link_floor_unit.Columns = new ColumnSet(new string[] { "bsd_block" });
            q.LinkEntities.Add(link_floor_unit);
            LinkEntity link_block_floor = new LinkEntity("bsd_floor", "bsd_block", "bsd_block", "bsd_blockid", JoinOperator.Inner);
            link_block_floor.EntityAlias = "bl";
            link_block_floor.Columns = new ColumnSet(new string[] { "bsd_project" });
            link_floor_unit.LinkEntities.Add(link_block_floor);
            LinkEntity link_project_block = new LinkEntity("bsd_block", "bsd_project", "bsd_project", "bsd_projectid", JoinOperator.Inner);
            link_project_block.EntityAlias = "pj";
            link_project_block.Columns = new ColumnSet(new string[] { "bsd_defaultpaymentscheme", "bsd_pricelistdefault" });
            link_block_floor.LinkEntities.Add(link_project_block);
            q.TopCount = 1;

            #region FetchXML

            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("");
            //sb.AppendLine("<fetch mapping='logical' count='1' output-format='xml-platform'>");
            //sb.AppendLine("<entity name='product'>");
            //sb.AppendLine("<attribute name='name'/>");
            //sb.AppendLine("<attribute name='productid'/>");
            //sb.AppendLine("<attribute name='bsd_floor'/>");
            //sb.AppendLine("<attribute name='bsd_blocknumber'/>");
            //sb.AppendLine("<attribute name='bsd_projectcode'/>");
            //sb.AppendLine("<attribute name='bsd_phaseslaunchid'/>");
            //sb.AppendLine("<attribute name='bsd_listprice'/>");
            //sb.AppendLine("<attribute name='defaultuomid'/>");
            //sb.AppendLine("<attribute name='statuscode'/>");
            //sb.AppendLine("<attribute name='statecode'/>");
            //sb.AppendLine("<filter type='and'>");
            //sb.AppendLine("<condition attribute='productid' operator='eq' value='" + unitId.ToString() + "'></condition>");
            //sb.AppendLine("</filter>");
            //sb.AppendLine("<link-entity name='bsd_floor' from='bsd_floorid' to='bsd_floor' link-type='inner'>");
            //sb.AppendLine("<attribute name='bsd_block'/>");
            //sb.AppendLine("<link-entity name='bsd_block' from='bsd_blockid' to='bsd_block' link-type='inner'>");
            //sb.AppendLine("<attribute name='bsd_project'/>");
            //sb.AppendLine("<link-entity name='bsd_project' from='bsd_projectid' to='bsd_project' link-type='inner'></link-entity>");
            //sb.AppendLine("</link-entity>");
            //sb.AppendLine("</link-entity>");
            //sb.AppendLine("</entity>");
            //sb.AppendLine("</fetch>"); 
            #endregion
            //EntityCollection entcs = service.RetrieveMultiple(new FetchExpression(sb.ToString()));
            EntityCollection entcs = service.RetrieveMultiple(q);
            if (entcs.Entities.Count == 0)
                return null;
            else
                return entcs.Entities[0];
        }

        private EntityReference PhasesLaunchPriceList(EntityReference enfPriceList)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_phaseslaunch"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_phaseslaunchid"" />
                <order attribute=""createdon"" descending=""true"" />
                <filter>
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000000"" />
                </filter>
                <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" intersect=""true"">
                  <filter>
                    <condition attribute=""bsd_pricelevelid"" operator=""eq"" value=""{enfPriceList.Id}"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection en = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (en.Entities.Count > 0)
                return en.Entities[0].ToEntityReference();
            else
                return null;
        }

        private Money GetQueuefee(EntityReference phaseF)
        {
            Money m = new Money(0);
            if (phaseF != null)
            {
                Entity tmp = service.Retrieve(phaseF.LogicalName, phaseF.Id, new ColumnSet(new string[] {
                    "bsd_bookingfee"
                }));
                if (tmp.Contains("bsd_bookingfee"))
                    m = (Money)tmp["bsd_bookingfee"];
            }
            return m;
        }

        private Money GetQepositfee(EntityReference pmSchRef)
        {
            Money money = new Money(Decimal.Zero);
            if (pmSchRef != null)
            {
                Entity entity = service.Retrieve(pmSchRef.LogicalName, pmSchRef.Id, new ColumnSet(new string[1]
                {
          "bsd_depositamount"
                }));
                if (entity.Contains("bsd_depositamount"))
                    money = (Money)entity["bsd_depositamount"];
            }
            return money;
        }
        private EntityCollection getListByIDCopy(IOrganizationService service, Guid idcopy)
        {
            #region --- Danh sách sắp xếp theo filter createdon mới nhất top 1 ---

            var fetchXml = $@"
<fetch>
  <entity name='pricelevel'>
    <all-attributes />
    <filter type='and'>
      <condition attribute='bsd_approved' operator='eq' value='1'/>
      <condition attribute='bsd_pricelistcopy' operator='eq' value='{idcopy}'/>
    </filter>
    <order attribute='createdon' descending='true' />
  </entity>
</fetch>";

            //var xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //      <entity name='pricelevel'>
            //        <attribute name='name' />
            //        <attribute name='transactioncurrencyid' />
            //        <attribute name='enddate' />
            //        <attribute name='begindate' />
            //        <attribute name='statecode' />
            //        <attribute name='createdon' />
            //        <attribute name='pricelevelid' />
            //        <order attribute='createdon' descending='true' />
            //        <filter type='and'>
            //          <condition attribute='bsd_approved' operator='eq' value='1' />
            //          <condition attribute='bsd_pricelistcopy' operator='eq'  uitype='pricelevel' value='" + idcopy + @"' />
            //        </filter>
            //      </entity>
            //    </fetch>";
            #endregion

            var rs = service.RetrieveMultiple(new FetchExpression(fetchXml.ToString()));
            return rs;
        }

        private EntityReference getBankAccount(Guid customerId)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_bankaccount"">
                <attribute name=""bsd_name"" />
                <filter>
                  <condition attribute=""bsd_customer"" operator=""eq"" value=""{customerId}"" />
                  <condition attribute=""bsd_default"" operator=""eq"" value=""1"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (result.Entities.Count <= 0) return null;
            return result.Entities[0].ToEntityReference();
        }

        [DataContract]
        public class InputParameter
        {
            [DataMember]
            public string action { get; set; }

            [DataMember]
            public string name { get; set; }

            [DataMember]
            public string value { get; set; }
        }
        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };
            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);

            return response.LocalTime;
            //var utcTime = utcTime.ToString("MM/dd/yyyy HH:mm:ss");
            //var localDateOnly = response.LocalTime.ToString("dd-MM-yyyy");
        }

        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            }).Entities[0].ToEntity<Entity>();
            return (int?)currentUserSettings.Attributes["timezonecode"];
        }
    }

}
