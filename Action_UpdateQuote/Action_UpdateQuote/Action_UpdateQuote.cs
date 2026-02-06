using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;

namespace Action_UpdateQuote
{
    public class Action_UpdateQuote : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("=== START Action_UpdateQuote ===");
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                tracingService.Trace($"Context Depth: {context.Depth}");
                string str1 = context.InputParameters["Command"].ToString();
                if (context.Depth > 3)
                {
                    tracingService.Trace("Depth > 3 => Stop plugin");
                    return;
                }

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                tracingService.Trace("Getting Target...");
                EntityReference target = (EntityReference)context.InputParameters["Target"];
                tracingService.Trace($"Target Entity: {target.LogicalName}, ID: {target.Id}");

                // Retrieve Quote
                tracingService.Trace("Retrieving Quote record...");
                Entity quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                tracingService.Trace("Quote retrieved successfully.");

                    Entity up_quote = new Entity(quote.LogicalName, quote.Id);
                    Entity entity_unit = service.Retrieve(((EntityReference)quote["bsd_unitno"]).LogicalName, ((EntityReference)quote["bsd_unitno"]).Id, new ColumnSet(true));
                    Entity up_unit = new Entity(entity_unit.LogicalName, entity_unit.Id);

                if (str1 == "confirm")
                {
                    // Fetch Tiến độ thanh toán
                    tracingService.Trace("Running FetchXML to get bsd_paymentschemedetail...");
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                    <fetch>
                      <entity name=""bsd_paymentschemedetail"">
                        <attribute name=""bsd_name"" />
                        <filter>
                          <condition attribute=""bsd_reservation"" operator=""eq"" value=""{target.Id}"" />
                          <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                        </filter>
                      </entity>
                    </fetch>";

                    tracingService.Trace("FetchXML: " + fetchXml);

                    EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));


                    if (rs.Entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException("Vui lòng tạo Tiến độ thanh toán");
                    }

                    // Update fields
                    up_quote["bsd_confirmdate"] = DateTime.Today;
                    up_quote["bsd_confirmer"] = new EntityReference("systemuser", context.UserId);
                    up_quote["statuscode"] = new OptionSetValue(100000000);
                    service.Update(up_quote);
                    up_unit["statuscode"] = new OptionSetValue(100000003);
                    service.Update(up_unit);
                    if (quote.Contains("bsd_opportunityid"))
                    {
                        Entity entity_booking = service.Retrieve(((EntityReference)quote["bsd_opportunityid"]).LogicalName, ((EntityReference)quote["bsd_opportunityid"]).Id, new ColumnSet(true));
                        Entity up_booking = new Entity(entity_booking.LogicalName, entity_booking.Id);
                    
                        var fetchXml_updatebooking = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_opportunity"">
                            <filter>
                              <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{entity_booking.Id}"" />
                              <condition attribute=""bsd_unit"" operator=""eq"" value=""{entity_unit.Id}"" />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection rs_booking = service.RetrieveMultiple(new FetchExpression(fetchXml_updatebooking));


                        if (rs_booking.Entities.Count > 0)
                        {
                            foreach (var entity in rs_booking.Entities)
                            {
                                Entity updateTarget = new Entity("bsd_opportunity", entity.Id);
                                updateTarget["statecode"] = new OptionSetValue(1);
                                updateTarget["statuscode"] = new OptionSetValue(100000005);//canceled
                                service.Update(updateTarget);
                            }
                        }
                    }
                    
                    tracingService.Trace("Quote updated successfully");
                    tracingService.Trace("Updating BPF Stage using Late-bound...");
                    tracingService.Trace("Updating BPF Stage via Process Instance...");

                    // 1. Tìm bản ghi Instance của quy trình đang chạy trên Quote này
                    // Tên thực thể BPF thường là tên Quy trình viết liền (ví dụ: bsd_bpf_deposit_process)
                    QueryExpression bpfQuery = new QueryExpression("new_reservationprocess") // Thay bằng Schema Name của BPF
                    {
                        ColumnSet = new ColumnSet("businessprocessflowinstanceid"),
                        Criteria = new FilterExpression()
                    };
                    bpfQuery.Criteria.AddCondition("bpf_bsd_quoteid", ConditionOperator.Equal, quote.Id);

                    EntityCollection bpfInstances = service.RetrieveMultiple(bpfQuery);

                    if (bpfInstances.Entities.Count > 0)
                    {
                        Entity bpfInstance = bpfInstances.Entities[0];

                        // 2. Cập nhật Stage hiện tại cho Instance này
                        bpfInstance["activestageid"] = new EntityReference("processstage", new Guid("8afac8a7-4a01-4e87-9d5a-700fc50b26f7"));

                        service.Update(bpfInstance);
                        tracingService.Trace("BPF Instance updated successfully.");
                    }
                    else
                    {
                        tracingService.Trace("No BPF Instance found for this Quote.");
                        // Nếu không tìm thấy, có nghĩa là bản ghi này chưa bao giờ được gán quy trình này
                    }
                    tracingService.Trace("BPF Stage updated successfully.");
                }
                if (str1 == "cancel")
                {
                    checkpayment(quote.Id, service);
                    if (quote.Contains("bsd_opportunityid")){
                        Entity queue = service.Retrieve(((EntityReference)quote["bsd_opportunityid"]).LogicalName, ((EntityReference)quote["bsd_opportunityid"]).Id, new ColumnSet(true));
                        Entity up_queue = new Entity(queue.LogicalName, queue.Id);
                        
                        tracingService.Trace("vào if cancel");
                        up_quote["statecode"] = new OptionSetValue(1);
                        up_quote["statuscode"] = new OptionSetValue(667980005);
                        up_quote["bsd_canceldate"] = DateTime.Today;
                        up_quote["bsd_canceller"] = new EntityReference("systemuser", context.UserId);
                        service.Update(up_quote);
                        
                        OptionSetValue statusCode = entity_unit.GetAttributeValue<OptionSetValue>("statuscode");
                        if (statusCode != null && statusCode.Value != 100000004)
                        {
                            up_unit["statuscode"] = new OptionSetValue(100000000);
                            service.Update(up_unit);
                        }
                        up_queue["statuscode"] = new OptionSetValue(100000004);
                        service.Update(up_queue);
                    }
                    else
                    {
                        tracingService.Trace("vào else cancel");
                        
                        up_quote["statuscode"] = new OptionSetValue(100000002);
                        up_quote["bsd_canceldate"] = DateTime.Today;
                        up_quote["bsd_canceller"] = new EntityReference("systemuser", context.UserId);
                        service.Update(up_quote);
                        OptionSetValue statusCode = entity_unit.GetAttributeValue<OptionSetValue>("statuscode");
                        if (statusCode != null && statusCode.Value != 100000004)
                        {
                            up_unit["statuscode"] = new OptionSetValue(100000000);
                            service.Update(up_unit);
                        }
                    }
                    
                }



            }
            catch (Exception ex)
            {
                // Trace toàn bộ lỗi
                tracingService.Trace("=== ERROR Action_UpdateQuote ===");
                tracingService.Trace("Message: " + ex.Message);
                tracingService.Trace("StackTrace: " + ex.StackTrace);

                // Ném lỗi ra ngoài để Action bắt được
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void checkpayment(Guid quoteId, IOrganizationService service)
        {
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_payment"">
                <filter>
                  <condition attribute=""statuscode"" operator=""in"">
                    <value>{1}</value>
                    <value>{100000000}</value>
                  </condition>
                  <condition attribute=""bsd_quotationreservation"" operator=""eq"" value=""{quoteId}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (rs.Entities.Count > 0)
            {
                throw new InvalidPluginExecutionException("The transaction has an invalid receipt. Please check again.");
            }
        }
        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");

            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };
            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
            return response.LocalTime;
        }

        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            }).Entities[0].ToEntity<Entity>();

            return (int?)currentUserSettings.Attributes["timezonecode"];
        }
    }
}
