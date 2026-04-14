using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Action_HandoverNotices_Generate
{
    public class Action_HandoverNotices_Generate : IPlugin
    {
        public static IOrganizationService service = null;
        static IOrganizationServiceFactory factory = null;
        ITracingService traceService = null;
        static StringBuilder strMess = new StringBuilder();
        static StringBuilder strMess2 = new StringBuilder();
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            string input01 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input01"]))
            {
                input01 = context.InputParameters["input01"].ToString();
            }
            string input02 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input02"]))
            {
                input02 = context.InputParameters["input02"].ToString();
            }
            string input03 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input03"]))
            {
                input03 = context.InputParameters["input03"].ToString();
            }
            string input04 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input04"]))
            {
                input04 = context.InputParameters["input04"].ToString();
            }
            if (input01 == "Buoc01" && input02 != "")
            {
                traceService.Trace("Bước 01");
                Entity enUp = new Entity("bsd_handovernotices");
                enUp.Id = Guid.Parse(input02);
                Entity enTarget = service.Retrieve(enUp.LogicalName, enUp.Id, new ColumnSet(true));
                int bsd_type = enTarget.Contains("bsd_type") ? ((OptionSetValue)enTarget["bsd_type"]).Value : 0;
                //LAY DANH SACH CAC SPA HOP LE
                var query = new QueryExpression("bsd_salesorder");
                query.ColumnSet.AddColumn("bsd_salesorderid");
                if (bsd_type == 100000000)//Property Handover
                    query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 100000013);
                else if (bsd_type == 100000001)//Certificate Handover
                {
                    query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 100000015);
                    query.Criteria.AddCondition("bsd_totalpercent", ConditionOperator.Equal, 100);
                    query.Criteria.AddCondition("bsd_totalinterestremaining", ConditionOperator.Equal, 0);
                }
                else if (bsd_type == 100000002)//Submit To Authority
                {
                    query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 100000013);
                }
                query.Criteria.AddCondition("bsd_project", ConditionOperator.Equal, ((EntityReference)enTarget["bsd_project"]).Id);
                if (bsd_type == 100000000)//Property Handover
                    query.Criteria.AddCondition("bsd_havehandovernotices", ConditionOperator.NotEqual, true);
                var U = query.AddLink("bsd_product", "bsd_unitnumber", "bsd_productid");
                U.EntityAlias = "U";
                if (enTarget.Contains("bsd_block")) U.LinkCriteria.AddCondition("bsd_blocknumber", ConditionOperator.Equal, ((EntityReference)enTarget["bsd_block"]).Id);
                if (enTarget.Contains("bsd_floor")) U.LinkCriteria.AddCondition("bsd_floor", ConditionOperator.Equal, ((EntityReference)enTarget["bsd_floor"]).Id);
                if (enTarget.Contains("bsd_unit")) U.LinkCriteria.AddCondition("bsd_productid", ConditionOperator.Equal, ((EntityReference)enTarget["bsd_unit"]).Id);
                var query_bsd_updateestimatehandoverdatedetail = query.AddLink(
                "bsd_updateestimatehandoverdatedetail",
                "bsd_salesorderid",
                "bsd_optionentry");
                var query_bsd_updateestimatehandoverdatedetail_bsd_updateestimatehandoverdate = query_bsd_updateestimatehandoverdatedetail.AddLink(
                    "bsd_updateestimatehandoverdate",
                    "bsd_updateestimatehandoverdate",
                    "bsd_updateestimatehandoverdateid");

                query_bsd_updateestimatehandoverdatedetail_bsd_updateestimatehandoverdate.LinkCriteria.AddCondition("bsd_createhandovernotices", ConditionOperator.Equal, true);
                EntityCollection list = service.RetrieveMultiple(query);
                List<string> listUnit = new List<string>();
                foreach (Entity detail in list.Entities)
                {
                    listUnit.Add(detail.Id.ToString());
                }
                List<string> listUnitGold = list.Entities
                   .Select(e => e.Id.ToString())
                   .Distinct()
                   .ToList();
                if (listUnitGold.Count == 0)
                    throw new InvalidPluginExecutionException("The list is empty. Please check again.");
                enUp["bsd_processing"] = true;
                enUp["bsd_list"] = string.Join(";", listUnitGold);
                service.Update(enUp);
            }
            else if (input01 == "Buoc02" && input02 != "" && input03 != "" && input04 != "")
            {
                traceService.Trace("Bước 02");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                Entity enTarget = service.Retrieve("bsd_handovernotices", Guid.Parse(input02), new ColumnSet(true));
                int bsd_type = enTarget.Contains("bsd_type") ? ((OptionSetValue)enTarget["bsd_type"]).Value : 0;
                Entity enSPA = service.Retrieve("bsd_salesorder", Guid.Parse(input03), new ColumnSet(
                    new string[] { "bsd_customerid", "bsd_project", "bsd_unitnumber", "bsd_totalpercent", "bsd_totalamountpaid", "bsd_depositfee", "bsd_freightamount", "bsd_numberofmonthspaidmf",
                        "bsd_managementfee", "bsd_totalinterest", "bsd_totalinterestpaid", "bsd_totalinterestremaining", "bsd_name" }));
                Entity enNew = new Entity("bsd_handovernoticedetail");
                enNew["bsd_name"] = "Notice Type - " + (string)enSPA["bsd_name"];
                enNew["bsd_handovernotices"] = enTarget.ToEntityReference();
                enNew["bsd_notificationattemptno"] = countHN(enSPA.Id, bsd_type);
                enNew["bsd_customer"] = enSPA.Contains("bsd_customerid") ? enSPA["bsd_customerid"] : null;
                enNew["bsd_project"] = enSPA.Contains("bsd_project") ? enSPA["bsd_project"] : null;
                enNew["bsd_unit"] = enSPA.Contains("bsd_unitnumber") ? enSPA["bsd_unitnumber"] : null;
                enNew["bsd_spa"] = enSPA.ToEntityReference();
                enNew["bsd_totalamountpaid"] = enSPA.Contains("bsd_totalpercent") ? enSPA["bsd_totalpercent"] : null;
                enNew["bsd_totalamountpaid2"] = enSPA.Contains("bsd_totalamountpaid") ? enSPA["bsd_totalamountpaid"] : null;
                enNew["bsd_depositamount"] = enSPA.Contains("bsd_depositfee") ? enSPA["bsd_depositfee"] : null;
                string nameField = bsd_type == 100000001 ? "bsd_lastinstallment" : "bsd_pinkbookhandover";
                var fetchXml_instalment = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                                        <fetch>
                                          <entity name=""bsd_paymentschemedetail"">
                                            <attribute name=""bsd_paymentschemedetailid"" />
                                            <attribute name=""bsd_amountofthisphase"" />
                                            <attribute name=""bsd_ordernumber"" />
                                            <attribute name=""bsd_duedate"" />
                                            <filter>
                                              <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{enSPA.Id}"" />
                                              <condition attribute=""{nameField}"" operator=""eq"" value=""1"" />
                                              <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                            </filter>
                                          </entity>
                                        </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml_instalment));
                decimal bsd_amountofthisphase = 0;
                decimal sumBalance = 0;
                foreach (Entity entity in rs.Entities)
                {
                    int bsd_ordernumber = (int)entity["bsd_ordernumber"];
                    enNew["bsd_installment"] = entity.ToEntityReference();
                    if (entity.Contains("bsd_duedate")) enNew["bsd_duedate"] = RetrieveLocalTimeFromUTCTime((DateTime)entity["bsd_duedate"], service);
                    bsd_amountofthisphase = entity.Contains("bsd_amountofthisphase") ? ((Money)entity["bsd_amountofthisphase"]).Value : 0;
                    enNew["bsd_installmentamount"] = new Money(bsd_amountofthisphase);
                    fetchXml_instalment = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                                        <fetch>
                                          <entity name=""bsd_paymentschemedetail"">
                                            <attribute name=""bsd_balance"" />
                                            <filter>
                                              <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{enSPA.Id}"" />
                                              <condition attribute=""bsd_ordernumber"" operator=""lt"" value=""{bsd_ordernumber}"" />
                                              <condition attribute=""bsd_balance"" operator=""gt"" value=""0"" />
                                              <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                            </filter>
                                          </entity>
                                        </fetch>";
                    EntityCollection rs2 = service.RetrieveMultiple(new FetchExpression(fetchXml_instalment));
                    foreach (Entity item in rs2.Entities)
                    {
                        sumBalance += item.Contains("bsd_balance") ? ((Money)item["bsd_balance"]).Value : 0;
                    }
                    enNew["bsd_outstandingunpaidinstallments"] = new Money(sumBalance);
                }
                decimal bsd_freightamount = enSPA.Contains("bsd_freightamount") ? ((Money)enSPA["bsd_freightamount"]).Value : 0;
                enNew["bsd_maintenancefee"] = new Money(bsd_freightamount);
                enNew["bsd_numberofmonthspaidmf"] = enSPA.Contains("bsd_numberofmonthspaidmf") ? enSPA["bsd_numberofmonthspaidmf"] : null;
                decimal bsd_managementfee = enSPA.Contains("bsd_managementfee") ? ((Money)enSPA["bsd_managementfee"]).Value : 0;
                enNew["bsd_managementfee"] = new Money(bsd_managementfee);
                enNew["bsd_totalinterestamount"] = enSPA.Contains("bsd_totalinterest") ? enSPA["bsd_totalinterest"] : null;
                enNew["bsd_totalinterestpaid"] = enSPA.Contains("bsd_totalinterestpaid") ? enSPA["bsd_totalinterestpaid"] : null;
                decimal bsd_totalinterestremaining = enSPA.Contains("bsd_totalinterestremaining") ? ((Money)enSPA["bsd_totalinterestremaining"]).Value : 0;
                enNew["bsd_totalinterestremaining"] = new Money(bsd_totalinterestremaining);
                enNew["bsd_totalamount"] = new Money((bsd_type == 100000001 ? 0 : bsd_freightamount + bsd_managementfee) + bsd_amountofthisphase + sumBalance + bsd_totalinterestremaining);
                service.Create(enNew);
                Entity enUpSPA = new Entity(enSPA.LogicalName, enSPA.Id);
                enUpSPA["bsd_havehandovernotices"] = true;
                service.Update(enUpSPA);
            }
            else if (input01 == "Buoc03" && input02 != "" && input04 != "")
            {
                traceService.Trace("Bước 03");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                Entity enUp = new Entity("bsd_handovernotices");
                enUp.Id = Guid.Parse(input02);
                enUp["bsd_processing"] = false;
                enUp["statuscode"] = new OptionSetValue(100000001);
                enUp["bsd_generateddate"] = DateTime.Now;
                enUp["bsd_generatedby"] = new EntityReference("systemuser", Guid.Parse(input04));
                service.Update(enUp);
            }
        }
        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
           new QueryExpression("usersettings")
           {
               ColumnSet = new ColumnSet("localeid", "timezonecode"),
               Criteria = new FilterExpression
               {
                   Conditions =
           {
            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
           }
               }
           }).Entities[0].ToEntity<Entity>();

            int? timeZoneCode = (int?)currentUserSettings.Attributes["timezonecode"];
            if (!timeZoneCode.HasValue)
                throw new Exception("Can't find time zone code");

            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);

            return response.LocalTime;
        }
        private int countHN(Guid inSPA, int bsd_type)
        {
            var fetchXml_instalment = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                                        <fetch>
                                          <entity name=""bsd_handovernoticedetail"">
                                            <attribute name=""bsd_handovernoticedetailid"" />
                                            <filter>
                                              <condition attribute=""bsd_spa"" operator=""eq"" value=""{inSPA}"" />
                                              <condition attribute=""statuscode"" operator=""eq"" value=""1"" />
                                            </filter>
                                            <link-entity name=""bsd_handovernotices"" from=""bsd_handovernoticesid"" to=""bsd_handovernotices"" alias=""HN"">
                                              <filter>
                                                <condition attribute=""bsd_type"" operator=""eq"" value=""{bsd_type}"" />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml_instalment));
            return rs.Entities.Count + 1;
        }
    }
}
