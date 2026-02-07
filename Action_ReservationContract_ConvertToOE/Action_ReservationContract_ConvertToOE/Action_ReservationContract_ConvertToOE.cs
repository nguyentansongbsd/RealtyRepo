using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_ReservationContract_ConvertToOE
{
    public class Action_ReservationContract_ConvertToOE : IPlugin
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

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity enRC = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_unitno", "bsd_projectid",
                "bsd_phaseslaunchid", "bsd_pricelevel", "bsd_paymentscheme", "bsd_handovercondition", "bsd_taxcode", "bsd_queuingfee", "bsd_depositfee",
                "bsd_netusablearea", "bsd_customerid", "bsd_bankaccount", "bsd_queue", "bsd_salessgentcompany", "bsd_detailamount", "bsd_discountamount",
                "bsd_packagesellingamount", "bsd_totalamountlessfreight", "bsd_totaltax", "bsd_totalamount", "bsd_quoteid", "bsd_discountcheck", "bsd_discountdraw",
                "bsd_freightamount", "bsd_numberofmonthspaidmf", "bsd_managementfee", "bsd_totalamountpaid", "bsd_totalpercent",
                "bsd_totalinterest", "bsd_totalinterestpaid", "bsd_totalinterestremaining", "bsd_customertype"}));
                int status = enRC.Contains("statuscode") ? ((OptionSetValue)enRC["statuscode"]).Value : -99;
                if (status != 100000000) //Active
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_reservationcontract"));

                if (!enRC.Contains("bsd_unitno"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_unitnumber"));
                EntityReference refProduct = (EntityReference)enRC["bsd_unitno"];
                Entity enProduct = service.Retrieve(refProduct.LogicalName, refProduct.Id, new ColumnSet(new string[] { "statuscode", "bsd_unittype" }));
                int statusProduct = enProduct.Contains("statuscode") ? ((OptionSetValue)enProduct["statuscode"]).Value : -99;
                if (statusProduct != 100000006) //Reserve
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_unit"));

                Guid idOE = CreateOE(enRC, target, refProduct, enProduct);
                EntityReference refOE = new EntityReference("bsd_salesorder", idOE);

                MapCoowner(target, refOE);
                MapPaymentSchemeDetail(target, refOE);
                MapPromotion(target, refOE);
                MapDiscountTransaction(target, refOE);
                MapPayment(target, refOE);
                UpdateReservationContract(target);
                UpdateUnit(refProduct);

                context.OutputParameters["id"] = idOE.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Guid CreateOE(Entity enRC, EntityReference target, EntityReference refProduct, Entity enProduct)
        {
            traceService.Trace("CreateOE");

            Entity newOE = new Entity("bsd_salesorder");
            newOE["bsd_name"] = refProduct.Name;
            newOE["bsd_date"] = DateTime.UtcNow;
            newOE["bsd_project"] = GetValidFieldValue(enRC, "bsd_projectid");
            newOE["bsd_phaseslaunch"] = GetValidFieldValue(enRC, "bsd_phaseslaunchid");
            newOE["bsd_pricelevel"] = GetValidFieldValue(enRC, "bsd_pricelevel");
            newOE["bsd_paymentscheme"] = GetValidFieldValue(enRC, "bsd_paymentscheme");
            newOE["bsd_handovercondition"] = GetValidFieldValue(enRC, "bsd_handovercondition");
            newOE["bsd_taxcode"] = GetValidFieldValue(enRC, "bsd_taxcode");
            newOE["bsd_queuingfee"] = GetValidFieldValue(enRC, "bsd_queuingfee");
            newOE["bsd_unittype"] = GetValidFieldValue(enProduct, "bsd_unittype");
            newOE["bsd_depositfee"] = GetValidFieldValue(enRC, "bsd_depositfee");
            newOE["bsd_unitnumber"] = GetValidFieldValue(enRC, "bsd_unitno");
            newOE["bsd_netusablearea"] = GetValidFieldValue(enRC, "bsd_netusablearea");
            newOE["bsd_customerid"] = GetValidFieldValue(enRC, "bsd_customerid");
            newOE["bsd_bankaccount"] = GetValidFieldValue(enRC, "bsd_bankaccount");
            newOE["bsd_opportunityid"] = GetValidFieldValue(enRC, "bsd_queue");
            newOE["bsd_quoteid"] = GetValidFieldValue(enRC, "bsd_quoteid"); ;
            newOE["bsd_reservationcontract"] = target;
            newOE["bsd_salesagentcompany"] = GetValidFieldValue(enRC, "bsd_salessgentcompany");

            newOE["bsd_detailamount"] = GetValidFieldValue(enRC, "bsd_detailamount");
            newOE["bsd_discount"] = GetValidFieldValue(enRC, "bsd_discountamount");
            newOE["bsd_packagesellingamount"] = GetValidFieldValue(enRC, "bsd_packagesellingamount");
            newOE["bsd_totalamountlessfreight"] = GetValidFieldValue(enRC, "bsd_totalamountlessfreight");
            newOE["bsd_totaltax"] = GetValidFieldValue(enRC, "bsd_totaltax");
            newOE["bsd_totalamount"] = GetValidFieldValue(enRC, "bsd_totalamount");
            newOE["bsd_freightamount"] = GetValidFieldValue(enRC, "bsd_freightamount");
            newOE["bsd_numberofmonthspaidmf"] = GetValidFieldValue(enRC, "bsd_numberofmonthspaidmf");
            newOE["bsd_managementfee"] = GetValidFieldValue(enRC, "bsd_managementfee");
            newOE["bsd_customertype"] = GetValidFieldValue(enRC, "bsd_customertype");

            newOE["bsd_discountcheck"] = GetValidFieldValue(enRC, "bsd_discountcheck");
            newOE["bsd_discountdraw"] = GetValidFieldValue(enRC, "bsd_discountdraw");

            newOE["bsd_totalamountpaid"] = GetValidFieldValue(enRC, "bsd_totalamountpaid");
            newOE["bsd_totalpercent"] = GetValidFieldValue(enRC, "bsd_totalpercent");
            newOE["bsd_totalinterest"] = GetValidFieldValue(enRC, "bsd_totalinterest");
            newOE["bsd_totalinterestpaid"] = GetValidFieldValue(enRC, "bsd_totalinterestpaid");
            newOE["bsd_totalinterestremaining"] = GetValidFieldValue(enRC, "bsd_totalinterestremaining");

            newOE.Id = Guid.NewGuid();
            service.Create(newOE);

            return newOE.Id;
        }

        private object GetValidFieldValue(Entity enRC, string field)
        {
            return enRC.Contains(field) ? enRC[field] : null; ;
        }

        private void MapCoowner(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapCoowner");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_coowner"">
                    <filter>
                      <condition attribute=""bsd_reservationcontract"" operator=""eq"" value=""{target.Id}"" />
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                    </filter>
                    <order attribute=""createdon"" />
                  </entity>
                </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, refOE);
                }
            }
        }

        private void MapPaymentSchemeDetail(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapPaymentSchemeDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_reservationcontract"" operator=""eq"" value=""{target.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, refOE);
                }
            }
        }

        private void CreateNewFromItem(Entity item, EntityReference refOE)
        {
            Entity it = new Entity(item.LogicalName);
            it = item;
            it.Attributes.Remove(item.LogicalName + "id");
            it.Attributes.Remove("ownerid");
            it.Attributes.Remove("bsd_reservationcontract");
            it["bsd_optionentry"] = refOE;
            it.Id = Guid.NewGuid();
            service.Create(it);
        }

        private void UpdateReservationContract(EntityReference target)
        {
            traceService.Trace("UpdateReservationContract");

            Entity upReservationContract = new Entity(target.LogicalName, target.Id);
            upReservationContract["statecode"] = new OptionSetValue(1);    //inactive
            upReservationContract["statuscode"] = new OptionSetValue(100000008);    //Convert to Option Entry
            service.Update(upReservationContract);
        }

        private void UpdateUnit(EntityReference refProduct)
        {
            traceService.Trace("UpdateUnit");

            Entity upUnit = new Entity(refProduct.LogicalName, refProduct.Id);
            upUnit["statuscode"] = new OptionSetValue(100000008);    //In Contract
            service.Update(upUnit);
        }

        private void MapPromotion(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapPromotion");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_promotion"">
                <attribute name=""bsd_promotionid"" />
                <attribute name=""bsd_name"" />
                <order attribute=""createdon"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
                <link-entity name=""bsd_bsd_reservationcontract_bsd_promotion"" from=""bsd_promotionid"" to=""bsd_promotionid"" intersect=""true"">
                  <filter>
                    <condition attribute=""bsd_reservationcontractid"" operator=""eq"" value=""{target.Id}"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                EntityReferenceCollection relativeEntity = new EntityReferenceCollection();
                foreach (var item in rs.Entities)
                {
                    relativeEntity.Add(new EntityReference(item.LogicalName, item.Id));
                }
                Relationship relationship = new Relationship("bsd_bsd_salesorder_bsd_promotion");
                service.Associate(refOE.LogicalName, refOE.Id, relationship, relativeEntity);
            }
        }

        private void MapDiscountTransaction(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapDiscountTransaction");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_discounttransaction"">
                <filter>
                  <condition attribute=""bsd_reservationcontract"" operator=""eq"" value=""{target.Id}"" />
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
                <order attribute=""createdon"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, refOE);
                }
            }
        }

        private void MapPayment(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapPayment");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_payment"">
                <filter>
                  <condition attribute=""bsd_reservationcontract"" operator=""eq"" value=""{target.Id}"" />
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
                <order attribute=""createdon"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, refOE);
                }
            }
        }
    }
}