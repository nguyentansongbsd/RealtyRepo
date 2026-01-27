using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Security.Principal;

namespace Action_QuotationReservation_ConvertToReservationContract
{
    public class Action_QuotationReservation_ConvertToReservationContract : IPlugin
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
                "bsd_packagesellingamount", "bsd_totalamountlessfreight", "bsd_vat", "bsd_totalamount","bsd_totalamountpaid", "bsd_discountcheck", "bsd_discountdraw"}));
                int status = enReservation.Contains("statuscode") ? ((OptionSetValue)enReservation["statuscode"]).Value : -99;
                
                EntityReference refProduct = (EntityReference)enReservation["bsd_unitno"];
                Entity enProduct = service.Retrieve(refProduct.LogicalName, refProduct.Id, new ColumnSet(new string[] { "statuscode", "bsd_unittype" }));
                int statusProduct = enProduct.Contains("statuscode") ? ((OptionSetValue)enProduct["statuscode"]).Value : -99;

                Guid idOE = CreateRAContract(enReservation, target, refProduct, enProduct);
                EntityReference refOE = new EntityReference("bsd_reservationcontract", idOE);

                //MapCoowner(target, refOE);
                //MapPaymentSchemeDetail(target, refOE);
                //MapPromotion(target, refOE);
                //MapDiscountTransaction(target, refOE);
                UpdateReservation(target);
                UpdateUnit(refProduct);

                context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + idOE.ToString() + "'}";
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Guid CreateRAContract(Entity enReservation, EntityReference target, EntityReference refProduct, Entity enProduct)
        {
            traceService.Trace("CreateRAContract");

            Entity newOE = new Entity("bsd_reservationcontract");
            newOE["bsd_name"] = refProduct.Name;
            newOE["statuscode"] = new OptionSetValue(1);
            newOE["bsd_projectid"] = GetValidFieldValue(enReservation, "bsd_projectid");
            newOE["bsd_totalamountpaid"] = GetValidFieldValue(enReservation, "bsd_totalamountpaid");
            newOE["bsd_phaseslaunchid"] = GetValidFieldValue(enReservation, "bsd_phaseslaunchid");
            newOE["bsd_pricelevel"] = GetValidFieldValue(enReservation, "bsd_pricelevel");
            newOE["bsd_paymentscheme"] = GetValidFieldValue(enReservation, "bsd_paymentscheme");
            newOE["bsd_handovercondition"] = GetValidFieldValue(enReservation, "bsd_handovercondition");
            newOE["bsd_taxcode"] = GetValidFieldValue(enReservation, "bsd_taxcode");
            newOE["bsd_queuingfee"] = GetValidFieldValue(enReservation, "bsd_bookingfee");
            newOE["bsd_unittype"] = GetValidFieldValue(enProduct, "bsd_unittype");
            newOE["bsd_depositfee"] = GetValidFieldValue(enReservation, "bsd_depositfee");
            newOE["bsd_unitno"] = GetValidFieldValue(enReservation, "bsd_unitno");
            newOE["bsd_netusablearea"] = GetValidFieldValue(enReservation, "bsd_netusablearea");
            newOE["bsd_customerid"] = GetValidFieldValue(enReservation, "bsd_customerid");
            newOE["bsd_bankaccount"] = GetValidFieldValue(enReservation, "bsd_bankaccount");
            newOE["bsd_quoteid"] = target;
            newOE["bsd_salessgentcompany"] = GetValidFieldValue(enReservation, "bsd_salessgentcompany");
            newOE["bsd_detailamount"] = GetValidFieldValue(enReservation, "bsd_detailamount");
            newOE["bsd_discountamount"] = GetValidFieldValue(enReservation, "bsd_discountamount");
            newOE["bsd_packagesellingamount"] = GetValidFieldValue(enReservation, "bsd_packagesellingamount");
            newOE["bsd_totalamountlessfreight"] = GetValidFieldValue(enReservation, "bsd_totalamountlessfreight");
            newOE["bsd_totaltax"] = GetValidFieldValue(enReservation, "bsd_vat");
            newOE["bsd_totalamount"] = GetValidFieldValue(enReservation, "bsd_totalamount");

            newOE["bsd_discountcheck"] = GetValidFieldValue(enReservation, "bsd_discountcheck");
            newOE["bsd_discountdraw"] = GetValidFieldValue(enReservation, "bsd_discountdraw");
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
            newOE["bsd_racontractsigndate"] = DateTime.Today;
            newOE["bsd_reservationnumber"] = "RSC-" + nextNumber.ToString("D8");
            newOE.Id = Guid.NewGuid();
            service.Create(newOE);

            return newOE.Id;
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
            upReservation["statuscode"] = new OptionSetValue(667980006);//Convert to RA Contract
            upReservation["statecode"] = new OptionSetValue(1);//inactive
            service.Update(upReservation);
        }

        private void UpdateUnit(EntityReference refProduct)
        {
            traceService.Trace("UpdateUnit");

            Entity upUnit = new Entity(refProduct.LogicalName, refProduct.Id);
            upUnit["statuscode"] = new OptionSetValue(100000006);    //Reserve
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
    }
}