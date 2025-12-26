using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PhasesLaunch_UnitsAssociate
{
    public class Plugin_PhasesLaunch_UnitsAssociate : IPlugin
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
                if (context.Depth > 1) return;
                traceService.Trace($"start {context.MessageName}");

                string relationship = context.InputParameters.Contains("Relationship") ? ((Relationship)context.InputParameters["Relationship"]).SchemaName : string.Empty;
                traceService.Trace(relationship);

                if (string.IsNullOrWhiteSpace(relationship) || !"bsd_bsd_phaseslaunch_bsd_pricelevel".Equals(relationship) ||
                        !(context.InputParameters.Contains("RelatedEntities") && context.InputParameters["RelatedEntities"] is EntityReferenceCollection))
                    return;

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                EntityReferenceCollection relatedEntities = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];
                traceService.Trace($"target {target.LogicalName} || {target.Id}");

                if (relatedEntities.Count == 0)
                    return;
                traceService.Trace($"relatedEntities {relatedEntities[0].LogicalName} || {relatedEntities[0].Id}");

                if ("Associate".Equals(context.MessageName))
                {
                    AssociateEntity(target, relatedEntities[0]);
                }
                else //Disassociate
                {
                    DisassociateEntity(relatedEntities[0], target);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void AssociateEntity(EntityReference refPriceList, EntityReference refPL)
        {
            traceService.Trace("AssociateEntity");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch distinct=""true"">
                      <entity name=""bsd_product"">
                        <attribute name=""bsd_productid"" />
                        <attribute name=""bsd_name"" />
                        <filter>
                          <condition entityname=""bsd_bsd_phaseslaunch_bsd_product"" attribute=""bsd_productid"" operator=""null"" />
                        </filter>
                        <order attribute=""bsd_name"" />
                        <link-entity name=""bsd_productpricelevel"" from=""bsd_product"" to=""bsd_productid"">
                          <filter>
                            <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{refPriceList.Id}"" />
                          </filter>
                        </link-entity>
                        <link-entity name=""bsd_bsd_phaseslaunch_bsd_product"" from=""bsd_productid"" to=""bsd_productid"" link-type=""outer"" alias=""bsd_bsd_phaseslaunch_bsd_product"" intersect=""true"">
                          <filter>
                            <condition attribute=""bsd_phaseslaunchid"" operator=""eq"" value=""{refPL.Id}"" />
                          </filter>
                        </link-entity>                      
                      </entity>
                    </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            traceService.Trace("rs.Entities.Count: " + rs.Entities.Count);
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                EntityReferenceCollection relativeEntity = new EntityReferenceCollection();
                foreach (var item in rs.Entities)
                {
                    relativeEntity.Add(new EntityReference(item.LogicalName, item.Id));
                }
                Relationship relationshipUnit = new Relationship("bsd_bsd_phaseslaunch_bsd_product");
                service.Associate(refPL.LogicalName, refPL.Id, relationshipUnit, relativeEntity);
            }
        }

        private void DisassociateEntity(EntityReference refPriceList, EntityReference refPL)
        {
            traceService.Trace("DisassociateEntity");

            List<Guid> listProduct = GetProductPriceList(refPriceList, refPL);
            if (listProduct.Count == 0)
                return;
            List<Guid> listProductPhasesLaunch = GetProductPhasesLaunch(refPL, listProduct);

            var productPhasesLaunchSet = new HashSet<Guid>(listProductPhasesLaunch);
            List<Guid> productRemove = listProduct
                .Where(id => !productPhasesLaunchSet.Contains(id))
                .ToList();

            traceService.Trace("productRemove: " + productRemove.Count);
            if (productRemove.Count > 0)
            {
                EntityReferenceCollection relativeEntity = new EntityReferenceCollection();
                foreach (var id in productRemove)
                {
                    relativeEntity.Add(new EntityReference("bsd_product", id));
                }
                Relationship relationshipUnit = new Relationship("bsd_bsd_phaseslaunch_bsd_product");
                service.Disassociate(refPL.LogicalName, refPL.Id, relationshipUnit, relativeEntity);
            }
        }

        private List<Guid> GetProductPriceList(EntityReference refPriceList, EntityReference refPL)
        {
            traceService.Trace("GetProductPriceList");
            List<Guid> tmpIds = new List<Guid>();

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch distinct=""true"">
              <entity name=""bsd_product"">
                <attribute name=""bsd_productid"" />
                <attribute name=""bsd_name"" />
                <link-entity name=""bsd_productpricelevel"" from=""bsd_product"" to=""bsd_productid"" alias=""bsd_productpricelevel"">
                    <filter>
                        <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{refPriceList.Id}"" />
                    </filter>
                </link-entity>
                <link-entity name=""bsd_bsd_phaseslaunch_bsd_product"" from=""bsd_productid"" to=""bsd_productid"" alias=""bsd_bsd_phaseslaunch_bsd_product"" intersect=""true"">
                  <filter>
                    <condition attribute=""bsd_phaseslaunchid"" operator=""eq"" value=""{refPL.Id}"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            traceService.Trace("rs.Entities.Count: " + rs.Entities.Count);
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                tmpIds = rs.Entities.Select(x => x.Id).ToList();
            }

            return tmpIds;
        }

        private List<Guid> GetProductPhasesLaunch(EntityReference refPL, List<Guid> listIdsRemove)
        {
            traceService.Trace("GetProductPhasesLaunch");
            List<Guid> tmpIds = new List<Guid>();

            StringBuilder tmp = new StringBuilder();
            tmp.AppendLine(@"<fetch distinct=""true"">");
            tmp.AppendLine(@"  <entity name=""bsd_product"">");
            tmp.AppendLine(@"    <attribute name=""bsd_productid"" />");
            tmp.AppendLine(@"    <attribute name=""bsd_name"" />");

            tmp.AppendLine(@"    <filter>");
            tmp.AppendLine(@"      <condition attribute=""bsd_productid"" operator=""in"">");
            foreach (var id in listIdsRemove)
            {
                tmp.AppendLine($@"        <value>{id}</value>");
            }
            tmp.AppendLine(@"      </condition>");
            tmp.AppendLine(@"    </filter>");

            tmp.AppendLine(@"    <link-entity name=""bsd_productpricelevel"" from=""bsd_product"" to=""bsd_productid"" alias=""bsd_productpricelevel"">");
            tmp.AppendLine(@"      <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevel"">");
            tmp.AppendLine(@"        <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">");
            tmp.AppendLine(@"          <filter>");
            tmp.AppendLine($@"            <condition attribute=""bsd_phaseslaunchid"" operator=""eq"" value=""{refPL.Id}"" />");
            tmp.AppendLine(@"          </filter>");
            tmp.AppendLine(@"        </link-entity>");
            tmp.AppendLine(@"      </link-entity>");
            tmp.AppendLine(@"    </link-entity>");
            tmp.AppendLine(@"    <link-entity name=""bsd_bsd_phaseslaunch_bsd_product"" from=""bsd_productid"" to=""bsd_productid"" alias=""bsd_bsd_phaseslaunch_bsd_product"" intersect=""true"" />");
            tmp.AppendLine(@"  </entity>");
            tmp.AppendLine(@"</fetch>");
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(tmp.ToString()));
            traceService.Trace("rs.Entities.Count: " + rs.Entities.Count);
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                tmpIds = rs.Entities.Select(x => x.Id).ToList();
            }

            return tmpIds;
        }
    }
}