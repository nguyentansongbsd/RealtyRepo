using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PriceList_Approved
{
    public class Plugin_PriceList_Approved : IPlugin
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
                Entity enPriceList = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode" }));
                int status = enPriceList.Contains("statuscode") ? ((OptionSetValue)enPriceList["statuscode"]).Value : -99;
                if (status != 100000000)  //Approved
                    return;

                if (CheckValidProduct(enPriceList))
                    throw new InvalidPluginExecutionException("The unit's status does not allow creating a price list.");

                UpPriceList(enPriceList);
                UpDetail(enPriceList);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private bool CheckValidProduct(Entity enPriceList)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_product"">
                <attribute name=""bsd_productid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""statuscode"" />
                <filter type=""or"">
                  <condition attribute=""statuscode"" operator=""not-in"">
                    <value>1</value>
                    <value>100000000</value>
                  </condition>
                  <condition attribute=""bsd_locked"" operator=""eq"" value=""1"" />
                </filter>
                <link-entity name=""bsd_productpricelevel"" from=""bsd_product"" to=""bsd_productid"" alias=""bsd_productpricelevel"">
                  <filter>
                    <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{enPriceList.Id}"" />
                    <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return (rs != null && rs.Entities != null && rs.Entities.Count > 0);
        }

        private void UpDetail(Entity enPriceList)
        {
            traceService.Trace("UpDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_productpricelevel"">
                <attribute name=""bsd_productpricelevelid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_product"" />
                <attribute name=""bsd_price"" />
                <attribute name=""bsd_usableareaunitprice"" />
                <attribute name=""bsd_builtupunitprice"" />
                <filter>
                    <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{enPriceList.Id}"" />
                    <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    Entity enUp = new Entity(item.LogicalName, item.Id);
                    enUp["statuscode"] = new OptionSetValue(100000000); //Approved
                    service.Update(enUp);

                    EntityReference refProduct = item.Contains("bsd_product") ? (EntityReference)item["bsd_product"] : null;
                    if (refProduct != null)
                        UpProduct(refProduct, enPriceList, item);
                }
            }
        }

        private void UpPriceList(Entity enPriceList)
        {
            traceService.Trace("UpPriceList");

            Entity upPriceList = new Entity(enPriceList.LogicalName, enPriceList.Id);
            upPriceList["statuscode"] = new OptionSetValue(100000000); //Approved
            service.Update(upPriceList);
        }

        private void UpProduct(EntityReference refProduct, Entity enPriceList, Entity item)
        {
            traceService.Trace("UpProduct");
            decimal bsd_price = item.Contains("bsd_price") ? ((Money)item["bsd_price"]).Value : 0;
            decimal bsd_usableareaunitprice = item.Contains("bsd_usableareaunitprice") ? ((Money)item["bsd_usableareaunitprice"]).Value : 0;
            decimal bsd_builtupunitprice = item.Contains("bsd_builtupunitprice") ? ((Money)item["bsd_builtupunitprice"]).Value : 0;

            Entity upProduct = new Entity(refProduct.LogicalName, refProduct.Id);
            upProduct["bsd_pricelevel"] = enPriceList.ToEntityReference();
            upProduct["bsd_price"] = new Money(bsd_price);
            upProduct["bsd_usableareaunitprice"] = new Money(bsd_usableareaunitprice);
            upProduct["bsd_builtupunitprice"] = new Money(bsd_builtupunitprice);
            upProduct["statuscode"] = new OptionSetValue(100000000); //Available
            service.Update(upProduct);
        }
    }
}