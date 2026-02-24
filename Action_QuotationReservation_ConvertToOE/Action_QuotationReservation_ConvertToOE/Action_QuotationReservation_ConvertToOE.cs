using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.IdentityModel.Metadata;
using System.Web.UI.WebControls;

namespace Action_QuotationReservation_ConvertToOE
{
    public class Action_QuotationReservation_ConvertToOE : IPlugin
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
                Entity enReservation = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_unitno", "bsd_projectid",
                "bsd_phaseslaunchid", "bsd_pricelevel", "bsd_paymentscheme", "bsd_handovercondition", "bsd_taxcode", "bsd_bookingfee", "bsd_depositfee",
                "bsd_netusablearea", "bsd_customerid", "bsd_bankaccount", "bsd_opportunityid", "bsd_salessgentcompany", "bsd_detailamount", "bsd_discountamount",
                "bsd_packagesellingamount", "bsd_totalamountlessfreight", "bsd_vat", "bsd_totalamount", "bsd_discountcheck", "bsd_discountdraw", "bsd_maintenancefees",
                "bsd_totalamountpaid", "bsd_customertype", "bsd_landvaluededuction", "bsd_numberofmonthspaidmf", "bsd_managementfee", "bsd_totalamountlessfreightaftervat"}));
                int status = enReservation.Contains("statuscode") ? ((OptionSetValue)enReservation["statuscode"]).Value : -99;
                if (status != 667980008) //Deposited
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_quotationreservation"));

                if (!enReservation.Contains("bsd_unitno"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_unitnumber"));
                EntityReference refProduct = (EntityReference)enReservation["bsd_unitno"];
                Entity enProduct = service.Retrieve(refProduct.LogicalName, refProduct.Id, new ColumnSet(new string[] { "statuscode", "bsd_unittype" }));
                int statusProduct = enProduct.Contains("statuscode") ? ((OptionSetValue)enProduct["statuscode"]).Value : -99;
                if (statusProduct != 100000003) //Deposited
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_unit"));

                Guid idOE = CreateOE(enReservation, target, refProduct, enProduct);
                EntityReference refOE = new EntityReference("bsd_salesorder", idOE);

                MapCoowner(target, refOE);
                MapPaymentSchemeDetail(target, refOE);
                MapPromotion(target, refOE);
                MapDiscountTransaction(target, refOE);
                //MapPayment(target, refOE);
                UpdateReservation(target);
                UpdateUnit(refProduct);

                context.OutputParameters["id"] = idOE.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Guid CreateOE(Entity enReservation, EntityReference target, EntityReference refProduct, Entity enProduct)
        {
            traceService.Trace("CreateOE");

            Entity newOE = new Entity("bsd_salesorder");
            newOE["bsd_name"] = refProduct.Name;
            newOE["bsd_date"] = DateTime.UtcNow;
            newOE["bsd_project"] = GetValidFieldValue(enReservation, "bsd_projectid");
            newOE["bsd_phaseslaunch"] = GetValidFieldValue(enReservation, "bsd_phaseslaunchid");
            newOE["bsd_pricelevel"] = GetValidFieldValue(enReservation, "bsd_pricelevel");
            newOE["bsd_paymentscheme"] = GetValidFieldValue(enReservation, "bsd_paymentscheme");
            newOE["bsd_handovercondition"] = GetValidFieldValue(enReservation, "bsd_handovercondition");
            newOE["bsd_taxcode"] = GetValidFieldValue(enReservation, "bsd_taxcode");
            newOE["bsd_queuingfee"] = GetValidFieldValue(enReservation, "bsd_bookingfee");
            newOE["bsd_unittype"] = GetValidFieldValue(enProduct, "bsd_unittype");
            newOE["bsd_depositfee"] = GetValidFieldValue(enReservation, "bsd_depositfee");
            newOE["bsd_unitnumber"] = GetValidFieldValue(enReservation, "bsd_unitno");
            newOE["bsd_netusablearea"] = GetValidFieldValue(enReservation, "bsd_netusablearea");
            newOE["bsd_customerid"] = GetValidFieldValue(enReservation, "bsd_customerid");
            newOE["bsd_bankaccount"] = GetValidFieldValue(enReservation, "bsd_bankaccount");
            newOE["bsd_opportunityid"] = GetValidFieldValue(enReservation, "bsd_opportunityid");
            newOE["bsd_quoteid"] = target;
            newOE["bsd_salesagentcompany"] = GetValidFieldValue(enReservation, "bsd_salessgentcompany");

            newOE["bsd_detailamount"] = GetValidFieldValue(enReservation, "bsd_detailamount");
            newOE["bsd_discount"] = GetValidFieldValue(enReservation, "bsd_discountamount");
            newOE["bsd_packagesellingamount"] = GetValidFieldValue(enReservation, "bsd_packagesellingamount");
            newOE["bsd_totalamountlessfreight"] = GetValidFieldValue(enReservation, "bsd_totalamountlessfreight");
            newOE["bsd_totalamountlessfreightaftervat"] = GetValidFieldValue(enReservation, "bsd_totalamountlessfreightaftervat");
            newOE["bsd_totaltax"] = GetValidFieldValue(enReservation, "bsd_vat");
            newOE["bsd_freightamount"] = GetValidFieldValue(enReservation, "bsd_maintenancefees");
            newOE["bsd_numberofmonthspaidmf"] = GetValidFieldValue(enReservation, "bsd_numberofmonthspaidmf");
            newOE["bsd_managementfee"] = GetValidFieldValue(enReservation, "bsd_managementfee");
            newOE["bsd_customertype"] = GetValidFieldValue(enReservation, "bsd_customertype");
            newOE["bsd_landvaluededuction"] = GetValidFieldValue(enReservation, "bsd_landvaluededuction");

            newOE["bsd_discountcheck"] = GetValidFieldValue(enReservation, "bsd_discountcheck");
            newOE["bsd_discountdraw"] = GetValidFieldValue(enReservation, "bsd_discountdraw");

            decimal bsd_totalamountpaid = enReservation.Contains("bsd_totalamountpaid") ? ((Money)enReservation["bsd_totalamountpaid"]).Value : 0;
            decimal bsd_totalamount = enReservation.Contains("bsd_totalamount") ? ((Money)enReservation["bsd_totalamount"]).Value : 0;
            newOE["bsd_totalamount"] = new Money(bsd_totalamount);
            newOE["bsd_totalamountpaid"] = new Money(bsd_totalamountpaid);

            newOE["bsd_totalpercent"] = bsd_totalamountpaid > 0 ? (bsd_totalamountpaid / bsd_totalamount * 100) : 0;

            newOE.Id = Guid.NewGuid();
            Guid id = service.Create(newOE);
            create_update_DataProjection(((EntityReference)GetValidFieldValue(enReservation, "bsd_unitno")).Id, newOE, id);
            return newOE.Id;
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
                enUp["bsd_spaid"] = new EntityReference("bsd_salesorder", id);
                if (enEntity.Contains("bsd_customerid")) enUp["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_project")) enUp["bsd_project"] = enEntity["bsd_project"];
                if (enEntity.Contains("bsd_opportunityid")) enUp["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_phaseslaunch")) enUp["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunch"];
                if (enEntity.Contains("bsd_reservationcontract")) enUp["bsd_raid"] = enEntity["bsd_reservationcontract"];
                if (enEntity.Contains("bsd_quoteid")) enUp["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Update(enUp);
            }
            else
            {
                Entity enCre = new Entity("bsd_dataprojection");
                enCre["bsd_spaid"] = new EntityReference("bsd_salesorder", id);
                if (enEntity.Contains("bsd_customerid")) enCre["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_project")) enCre["bsd_project"] = enEntity["bsd_project"];
                if (enEntity.Contains("bsd_opportunityid")) enCre["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_phaseslaunch")) enCre["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunch"];
                if (enEntity.Contains("bsd_unitno")) enCre["bsd_productid"] = enEntity["bsd_unitno"];
                if (enEntity.Contains("bsd_reservationcontract")) enCre["bsd_raid"] = enEntity["bsd_reservationcontract"];
                if (enEntity.Contains("bsd_quoteid")) enCre["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Create(enCre);
            }
        }
        private object GetValidFieldValue(Entity enReservation, string field)
        {
            return enReservation.Contains(field) ? enReservation[field] : null; ;
        }

        private void MapCoowner(EntityReference target, EntityReference refOE)
        {
            traceService.Trace("MapCoowner");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_coowner"">
                    <filter>
                      <condition attribute=""bsd_reservation"" operator=""eq"" value=""{target.Id}"" />
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
                    CreateNewFromItem(item, "bsd_reservation", refOE);
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
                  <condition attribute=""bsd_reservation"" operator=""eq"" value=""{target.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, "bsd_reservation", refOE);
                }
            }
        }

        private void CreateNewFromItem(Entity item, string logicalField, EntityReference refOE)
        {
            Entity it = new Entity(item.LogicalName);
            it = item;
            it.Attributes.Remove(item.LogicalName + "id");
            it.Attributes.Remove("ownerid");
            it.Attributes.Remove(logicalField);
            it["bsd_optionentry"] = refOE;
            it.Id = Guid.NewGuid();
            service.Create(it);
        }

        private void UpdateReservation(EntityReference target)
        {
            traceService.Trace("UpdateReservation");

            Entity upReservation = new Entity(target.LogicalName, target.Id);
            upReservation["statecode"] = new OptionSetValue(1);    //inactive
            upReservation["statuscode"] = new OptionSetValue(667980007);    //Convert to Option Entry
            service.Update(upReservation);
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
                <link-entity name=""bsd_bsd_quote_bsd_promotion"" from=""bsd_promotionid"" to=""bsd_promotionid"" intersect=""true"">
                  <filter>
                    <condition attribute=""bsd_quoteid"" operator=""eq"" value=""{target.Id}"" />
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
                  <condition attribute=""bsd_quote"" operator=""eq"" value=""{target.Id}"" />
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
                    CreateNewFromItem(item, "bsd_quote", refOE);
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
                  <condition attribute=""bsd_quotationreservation"" operator=""eq"" value=""{target.Id}"" />
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
                    CreateNewFromItem(item, "bsd_quotationreservation", refOE);
                }
            }
        }
    }
}