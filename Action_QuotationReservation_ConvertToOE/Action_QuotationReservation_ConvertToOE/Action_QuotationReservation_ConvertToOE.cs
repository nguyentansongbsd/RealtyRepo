using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                "bsd_packagesellingamount", "bsd_totalamountlessfreight", "bsd_vat", "bsd_totalamount"}));
                int status = enReservation.Contains("statuscode") ? ((OptionSetValue)enReservation["statuscode"]).Value : -99;
                if (status != 667980002) //Director Approval
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_quotationreservation"));

                if (!enReservation.Contains("bsd_unitno"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_unitnumber"));
                EntityReference refProduct = (EntityReference)enReservation["bsd_unitno"];
                Entity enProduct = service.Retrieve(refProduct.LogicalName, refProduct.Id, new ColumnSet(new string[] { "statuscode", "bsd_unittype" }));
                int statusProduct = enProduct.Contains("statuscode") ? ((OptionSetValue)enProduct["statuscode"]).Value : -99;
                if (statusProduct != 100000003) //Deposited
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_unit"));

                Guid idOE = CreateOE(enReservation, target, refProduct, enProduct);
                MapCoowner(target, idOE);
                MapPaymentSchemeDetail(target, idOE);
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
            newOE["bsd_totaltax"] = GetValidFieldValue(enReservation, "bsd_vat");
            newOE["bsd_totalamount"] = GetValidFieldValue(enReservation, "bsd_totalamount");

            newOE.Id = Guid.NewGuid();
            service.Create(newOE);

            return newOE.Id;
        }

        private object GetValidFieldValue(Entity enReservation, string field)
        {
            return enReservation.Contains(field) ? enReservation[field] : null; ;
        }

        private void MapCoowner(EntityReference target, Guid id)
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
                    CreateNewFromItem(item, id);
                }
            }
        }

        private void MapPaymentSchemeDetail(EntityReference target, Guid id)
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
                    CreateNewFromItem(item, id);
                }
            }
        }

        private void CreateNewFromItem(Entity item, Guid id)
        {
            Entity it = new Entity(item.LogicalName);
            it = item;
            it.Attributes.Remove(item.LogicalName + "id");
            it.Attributes.Remove("ownerid");
            it.Attributes.Remove("bsd_reservation");
            it["bsd_optionentry"] = new EntityReference("bsd_salesorder", id);
            it.Id = Guid.NewGuid();
            service.Create(it);
        }

        private void UpdateReservation(EntityReference target)
        {
            traceService.Trace("UpdateReservation");

            Entity upReservation = new Entity(target.LogicalName, target.Id);
            upReservation["statuscode"] = new OptionSetValue(100000012);    //Convert to Option Entry
            service.Update(upReservation);
        }

        private void UpdateUnit(EntityReference refProduct)
        {
            traceService.Trace("UpdateUnit");

            Entity upUnit = new Entity(refProduct.LogicalName, refProduct.Id);
            upUnit["statuscode"] = new OptionSetValue(100000008);    //In Contract
            service.Update(upUnit);
        }
    }
}