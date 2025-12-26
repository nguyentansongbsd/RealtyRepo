using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_UpdatePriceList_CreateUpdate
{
    public class Plugin_UpdatePriceList_CreateUpdate : IPlugin
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
                if (context.Depth > 2) return;

                Entity target = (Entity)context.InputParameters["Target"];
                Entity enUpdatePriceList = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_product", "bsd_pricelist", "bsd_usableareaunitpricenew" }));

                if (!enUpdatePriceList.Contains("bsd_pricelist"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Price list'. Vui lòng kiểm tra lại.");

                if (!enUpdatePriceList.Contains("bsd_product"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Product'. Vui lòng kiểm tra lại.");

                EntityReference refProduct = (EntityReference)enUpdatePriceList["bsd_product"];
                if ("Create".Equals(context.MessageName) && CheckValidProduct(refProduct))
                    throw new InvalidPluginExecutionException("The unit’s status does not permit updating the price list.");

                decimal bsd_usableareaunitpricenew = enUpdatePriceList.Contains("bsd_usableareaunitpricenew") ? ((Money)enUpdatePriceList["bsd_usableareaunitpricenew"]).Value : 0;

                EntityReference refPriceList = (EntityReference)enUpdatePriceList["bsd_pricelist"];
                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch top=""1"">
                  <entity name=""bsd_productpricelevel"">
                    <attribute name=""bsd_productpricelevelid"" />
                    <attribute name=""bsd_name"" />
                    <attribute name=""bsd_netusablearea"" />
                    <attribute name=""bsd_builtuparea"" />
                    <filter>
                        <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{refPriceList.Id}"" />
                        <condition attribute=""bsd_product"" operator=""eq"" value=""{refProduct.Id}"" />
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    foreach (var item in rs.Entities)
                    {
                        decimal bsd_netusablearea = item.Contains("bsd_netusablearea") ? (decimal)item["bsd_netusablearea"] : 0;

                        Entity enUp = new Entity(enUpdatePriceList.LogicalName, enUpdatePriceList.Id);
                        decimal bsd_pricenew = bsd_netusablearea * bsd_usableareaunitpricenew;
                        enUp["bsd_pricenew"] = new Money(bsd_pricenew);

                        decimal bsd_builtuparea = item.Contains("bsd_builtuparea") ? (decimal)item["bsd_builtuparea"] : 0;
                        decimal bsd_builtupunitprice = bsd_builtuparea != 0 ? bsd_pricenew / bsd_builtuparea : 0;
                        enUp["bsd_builtupunitpricenew"] = new Money(bsd_builtupunitprice);
                        service.Update(enUp);
                    }
                }

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
    }
}