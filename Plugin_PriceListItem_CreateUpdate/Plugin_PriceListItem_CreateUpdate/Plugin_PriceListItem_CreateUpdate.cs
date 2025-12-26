using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PriceListItem_CreateUpdate
{
    public class Plugin_PriceListItem_CreateUpdate : IPlugin
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
                Entity enPriceListItem = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_product", "bsd_usableareaunitprice" }));
                if (!enPriceListItem.Contains("bsd_product"))
                    throw new InvalidPluginExecutionException("Không có dữ liệu 'Product'. Vui lòng kiểm tra lại.");

                EntityReference refProduct = (EntityReference)enPriceListItem["bsd_product"];
                if ("Create".Equals(context.MessageName) && CheckValidProduct(refProduct))
                    throw new InvalidPluginExecutionException("The unit's status does not allow creating a price list.");

                decimal bsd_usableareaunitprice = enPriceListItem.Contains("bsd_usableareaunitprice") ? ((Money)enPriceListItem["bsd_usableareaunitprice"]).Value : 0;

                Entity enProduct = service.Retrieve(refProduct.LogicalName, refProduct.Id, new ColumnSet(new string[] { "bsd_netsaleablearea", "bsd_builtuparea" }));
                decimal bsd_netsaleablearea = enProduct.Contains("bsd_netsaleablearea") ? (decimal)enProduct["bsd_netsaleablearea"] : 0;
                decimal bsd_builtuparea = enProduct.Contains("bsd_builtuparea") ? (decimal)enProduct["bsd_builtuparea"] : 0;

                Entity upPriceListItem = new Entity(enPriceListItem.LogicalName, enPriceListItem.Id);
                upPriceListItem["bsd_netusablearea"] = bsd_netsaleablearea;
                decimal bsd_price = bsd_netsaleablearea * bsd_usableareaunitprice;
                upPriceListItem["bsd_price"] = new Money(bsd_price);
                upPriceListItem["bsd_builtuparea"] = bsd_builtuparea;
                decimal bsd_builtupunitprice = bsd_builtuparea != 0 ? bsd_price / bsd_builtuparea : 0;
                upPriceListItem["bsd_builtupunitprice"] = new Money(bsd_builtupunitprice);
                service.Update(upPriceListItem);

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
                <attribute name=""statuscode"" />
                <filter type=""and"">
                  <condition attribute=""bsd_productid"" operator=""eq"" value=""{refProduct.Id}"" />
                </filter>
                <filter type=""or"">
                  <condition attribute=""statuscode"" operator=""not-in"">
                    <value>1</value>
                    <value>100000000</value>
                  </condition>
                  <condition attribute=""bsd_locked"" operator=""eq"" value=""1"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return (rs != null && rs.Entities != null && rs.Entities.Count > 0);
        }
    }
}