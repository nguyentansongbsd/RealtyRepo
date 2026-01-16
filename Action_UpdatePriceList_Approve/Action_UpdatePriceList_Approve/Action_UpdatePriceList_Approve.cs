using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_UpdatePriceList_Approve
{
    public class Action_UpdatePriceList_Approve : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");

                EntityReference refUpdatePriceList = (EntityReference)context.InputParameters["Target"];
                Entity enUpdatePriceList = service.Retrieve(refUpdatePriceList.LogicalName, refUpdatePriceList.Id, new ColumnSet(new string[] { "bsd_product", "bsd_pricelist",
                                            "bsd_usableareaunitpricenew", "bsd_pricenew", "bsd_builtupunitpricenew", "bsd_powerautomate", "bsd_paprocess", "bsd_productpricelevel"}));

                if (!enUpdatePriceList.Contains("bsd_pricelist"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Price list'. Vui lòng kiểm tra lại.");

                if (!enUpdatePriceList.Contains("bsd_product"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Product'. Vui lòng kiểm tra lại.");

                if (!enUpdatePriceList.Contains("bsd_productpricelevel"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Price List Item'. Vui lòng kiểm tra lại.");

                EntityReference refProduct = (EntityReference)enUpdatePriceList["bsd_product"];
                EntityReference refPriceList = (EntityReference)enUpdatePriceList["bsd_pricelist"];
                EntityReference refPLI = (EntityReference)enUpdatePriceList["bsd_productpricelevel"];

                if (CheckValidProduct(refProduct))
                    throw new InvalidPluginExecutionException("The unit’s status does not permit updating the price list.");

                bool isPA = enUpdatePriceList.Contains("bsd_powerautomate") ? (bool)enUpdatePriceList["bsd_powerautomate"] : false;
                string paProcess = context.InputParameters.Contains("paProcess") && !string.IsNullOrEmpty((string)context.InputParameters["paProcess"]) ?
                                            (string)context.InputParameters["paProcess"] : string.Empty;
                Guid userId = context.InputParameters.Contains("userId") && !string.IsNullOrEmpty((string)context.InputParameters["userId"]) ?
                        Guid.Parse((string)context.InputParameters["userId"]) : context.UserId;
                traceService.Trace($"userId: {userId} || {paProcess}");

                if (isPA && enUpdatePriceList.Contains("bsd_paprocess") && (string)enUpdatePriceList["bsd_paprocess"] != paProcess)
                    throw new InvalidPluginExecutionException("Record này đang được thực hiện ở tiến trình khác. Vui lòng kiểm tra lại.");

                UpPriceListItem(refPLI, enUpdatePriceList);
                UpProduct(refProduct, refPriceList, enUpdatePriceList);
                UpUpdatePriceList(enUpdatePriceList, userId);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private bool CheckValidProduct(EntityReference refProduct)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_product"">
                <attribute name=""bsd_productid"" />
                <attribute name=""bsd_name"" />
                <filter type=""and"">
                  <condition attribute=""bsd_productid"" operator=""eq"" value=""{refProduct.Id}"" />
                </filter>
                <filter type=""or"">
                  <condition attribute=""statuscode"" operator=""ne"" value=""100000000"" />
                  <condition attribute=""bsd_locked"" operator=""eq"" value=""1"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return (rs != null && rs.Entities != null && rs.Entities.Count > 0);
        }

        private void UpUpdatePriceList(Entity enUpdatePriceList, Guid userId)
        {
            traceService.Trace("UpUpdatePriceList");
            Entity enUp = new Entity(enUpdatePriceList.LogicalName, enUpdatePriceList.Id);
            enUp["statuscode"] = new OptionSetValue(100000000); //Approve
            enUp["bsd_powerautomate"] = false;
            enUp["bsd_error"] = null;
            enUp["bsd_paprocess"] = null;
            enUp["bsd_approvaldate"] = DateTime.UtcNow;
            enUp["bsd_approver"] = new EntityReference("systemuser", userId);
            service.Update(enUp);
        }

        private void UpPriceListItem(EntityReference refPLI, Entity enUpdatePriceList)
        {
            traceService.Trace("UpPriceListItem");

            Entity enUp = new Entity(refPLI.LogicalName, refPLI.Id);
            enUp["bsd_usableareaunitprice"] = enUpdatePriceList.Contains("bsd_usableareaunitpricenew") ? enUpdatePriceList["bsd_usableareaunitpricenew"] : null;
            enUp["bsd_price"] = enUpdatePriceList.Contains("bsd_pricenew") ? enUpdatePriceList["bsd_pricenew"] : null;
            enUp["bsd_builtupunitprice"] = enUpdatePriceList.Contains("bsd_builtupunitpricenew") ? enUpdatePriceList["bsd_builtupunitpricenew"] : null;
            service.Update(enUp);
        }

        private void UpProduct(EntityReference refProduct, EntityReference refPriceList, Entity enUpdatePriceList)
        {
            traceService.Trace("UpProduct");

            Entity enUp = new Entity(refProduct.LogicalName, refProduct.Id);
            enUp["bsd_usableareaunitprice"] = enUpdatePriceList.Contains("bsd_usableareaunitpricenew") ? enUpdatePriceList["bsd_usableareaunitpricenew"] : null;
            enUp["bsd_price"] = enUpdatePriceList.Contains("bsd_pricenew") ? enUpdatePriceList["bsd_pricenew"] : null;
            enUp["bsd_builtupunitprice"] = enUpdatePriceList.Contains("bsd_builtupunitpricenew") ? enUpdatePriceList["bsd_builtupunitpricenew"] : null;
            enUp["bsd_pricelevel"] = refPriceList;
            service.Update(enUp);
        }
    }
}