using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Action_Segment_Dynamic
{
    public class Action_Segment_Dynamic : IPlugin
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
                traceService.Trace($"start {context.MessageName}");
                if (context.Depth > 2) return;

                EntityReference refSegment = (EntityReference)context.InputParameters["Target"];
                Entity enSegment = service.Retrieve(refSegment.LogicalName, refSegment.Id, new ColumnSet(new string[] { "bsd_type", "bsd_entitytype", "bsd_query" }));
                int bsd_type = enSegment.Contains("bsd_type") ? ((OptionSetValue)enSegment["bsd_type"]).Value : -99;
                if (bsd_type != 100000001 || !enSegment.Contains("bsd_query"))  //Dynamic
                    return;

                int bsd_entitytype = enSegment.Contains("bsd_entitytype") ? ((OptionSetValue)enSegment["bsd_entitytype"]).Value : -99;
                if (bsd_entitytype != 100000000)  //Contact
                    return;

                HashSet<Guid> customerIds = new HashSet<Guid>();
                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_querybuildergroup"">
                    <attribute name=""bsd_fetchxml"" />
                    <attribute name=""bsd_parententity"" />
                    <filter>
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                      <condition attribute=""bsd_regardingobjectid"" operator=""eq"" value=""{refSegment.Id}"" />
                      <condition attribute=""bsd_fetchxml"" operator=""not-null"" />
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    foreach (var item in rs.Entities)
                    {
                        GetListSegmentMember((string)item["bsd_fetchxml"], ref customerIds);
                    }
                }

                CreateSegmentMember(refSegment, customerIds);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void GetListSegmentMember(string bsd_fetchxml, ref HashSet<Guid> customerIds)
        {
            traceService.Trace("GetListSegmentMember");

            string field = "bsd_customerid";

            string newFetchXML = AddField(bsd_fetchxml, field);
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(newFetchXML));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    if (!item.Contains(field))
                        continue;

                    EntityReference refCustomer = (EntityReference)item[field];
                    if (refCustomer.LogicalName != "contact")
                        continue;

                    customerIds.Add(refCustomer.Id);
                }
            }
        }

        private void CreateSegmentMember(EntityReference refSegment, HashSet<Guid> customerIds)
        {
            traceService.Trace("CreateSegmentMember");

            EntityCollection listExist = GetExistSegmentMember(refSegment);
            Dictionary<Guid, Entity> existMap = new Dictionary<Guid, Entity>();
            foreach (var e in listExist.Entities)
            {
                EntityReference refCus = (EntityReference)e["bsd_customer"];
                existMap[refCus.Id] = e;
            }

            // Skip nếu không thay đổi
            if (customerIds.SetEquals(existMap.Keys))
            {
                traceService.Trace("No change → skip");
                return;
            }

            List<Entity> createList = new List<Entity>();
            List<EntityReference> deleteList = new List<EntityReference>();

            // create
            foreach (var id in customerIds)
            {
                if (!existMap.ContainsKey(id))
                {
                    Entity newSegmentMember = new Entity("bsd_segmentmember");
                    newSegmentMember["bsd_entitytype"] = new OptionSetValue(100000000);  //Contact
                    newSegmentMember["bsd_customer"] = new EntityReference("contact", id);
                    newSegmentMember["bsd_addeddate"] = DateTime.UtcNow;
                    newSegmentMember["bsd_segment"] = refSegment;
                    createList.Add(newSegmentMember);
                }
            }

            // delete
            foreach (var kv in existMap)
            {
                if (!customerIds.Contains(kv.Key))
                {
                    deleteList.Add(new EntityReference(kv.Value.LogicalName, kv.Value.Id));
                }
            }

            if (createList.Count > 0)
                BulkCreate(createList);

            if (deleteList.Count > 0)
                BulkDelete(deleteList);
        }

        private string AddField(string bsd_fetchxml, string field)
        {
            traceService.Trace("AddField");

            XDocument doc = XDocument.Parse(bsd_fetchxml);

            // tìm entity chính
            XElement entity = doc.Descendants("entity").FirstOrDefault();

            if (entity != null)
            {
                // check nếu chưa có thì mới add
                bool exists = entity.Elements("attribute").Any(a => a.Attribute("name")?.Value == field);
                if (!exists)
                {
                    entity.AddFirst(new XElement("attribute", new XAttribute("name", field)));
                }
            }

            return doc.ToString();
        }

        private EntityCollection GetExistSegmentMember(EntityReference refSegment)
        {
            traceService.Trace("GetExistSegmentMember");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_segmentmember"">
                <attribute name=""bsd_segmentmemberid"" />
                <attribute name=""bsd_customer"" />
                <filter>
                  <condition attribute=""bsd_segment"" operator=""eq"" value=""{refSegment.Id}"" />
                  <condition attribute=""bsd_customer"" operator=""not-null"" />
                </filter>
              </entity>
            </fetch>";
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        private void BulkCreate(List<Entity> entities)
        {
            traceService.Trace("BulkCreate");

            try
            {
                traceService.Trace($"requests {entities.Count}");

                ExecuteTransactionRequest transactionRequest = new ExecuteTransactionRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    ReturnResponses = false
                };

                foreach (var entity in entities)
                {
                    transactionRequest.Requests.Add(new CreateRequest { Target = entity });
                }

                service.Execute(transactionRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Transaction failed: {ex.Message}", ex);
            }
        }

        private void BulkDelete(List<EntityReference> entityReferences)
        {
            traceService.Trace("BulkDelete");

            try
            {
                traceService.Trace($"requests {entityReferences.Count}");

                ExecuteTransactionRequest transactionRequest = new ExecuteTransactionRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    ReturnResponses = false
                };

                foreach (var entityRef in entityReferences)
                {
                    transactionRequest.Requests.Add(new DeleteRequest { Target = entityRef });
                }

                service.Execute(transactionRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Transaction failed: {ex.Message}", ex);
            }
        }
    }
}