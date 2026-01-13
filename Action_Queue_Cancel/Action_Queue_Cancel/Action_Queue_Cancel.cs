using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Queue_Cancel
{
    public class Action_Queue_Cancel : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationServiceFactory _serviceFactory;
        private IOrganizationService _service;
        private ITracingService _tracingService;

        EntityReference Target = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            this.Target = this._context.InputParameters["Target"] as EntityReference;
            Entity enQueue = this._service.Retrieve(this.Target.LogicalName, this.Target.Id, new ColumnSet("bsd_unit", "bsd_collectedqueuingfee", "bsd_project", "bsd_customerid", "bsd_douutien", "bsd_queuingfeepaid", "bsd_name"));
            CancelQueue(this.Target);
            if (enQueue.Contains("bsd_unit"))
                UpdateUnit((EntityReference)enQueue["bsd_unit"]);
            if (enQueue.Contains("bsd_queuingfeepaid"))
                create_Refund(enQueue);
            UpdatePriority(enQueue);
        }
        private void CancelQueue(EntityReference target)
        {
            try
            {
                _tracingService.Trace("Start Cancel Queue");
                Entity queue = new Entity(target.LogicalName, target.Id);
                queue["statecode"] = new OptionSetValue(1); // Canceled
                queue["statuscode"] = new OptionSetValue(100000005); // Canceled
                queue["bsd_canceller"] = new EntityReference("systemuser", this._context.UserId);
                queue["bsd_cancelleddate"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, this._service);
                this._service.Update(queue);
                _tracingService.Trace("End Cancel Queue");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in the Action_Queue_Cancel plugin.", ex);
            }
        }
        // Update Unit Status to Available if no Opportunities queuing or waiting
        private void UpdateUnit(EntityReference enfUnit)
        {
            // 100000004 = Queuing, 100000003 = Waiting
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <filter type=""and"">
                  <condition attribute=""bsd_unit"" operator=""eq"" value=""{enfUnit.Id}"" />
                  <condition attribute=""statuscode"" operator=""in"">
                    <value>100000004</value>
                    <value>100000003</value>
                  </condition>
                </filter>
              </entity>
            </fetch>";
            EntityCollection ecOpportunities = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (ecOpportunities.Entities.Count == 0)
            {
                Entity enUnit = new Entity(enfUnit.LogicalName, enfUnit.Id);
                enUnit["statuscode"] = new OptionSetValue(100000000); // Available
                this._service.Update(enUnit);
            }
        }

        private void UpdatePriority(Entity enQueue)
        {
            if (enQueue.Contains("bsd_collectedqueuingfee") && (bool)enQueue["bsd_collectedqueuingfee"] == false) return;
            string conditionUnit = enQueue.Contains("bsd_unit") ? $@"<condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_unit"]).Id}"" />" : $@"<condition attribute=""bsd_unit"" operator=""null"" />";
            string conditionProject = enQueue.Contains("bsd_project") ? $@"<condition attribute=""bsd_project"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_project"]).Id}"" />" : "";
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_dateorder"" />
                <attribute name=""bsd_douutien"" />
                <filter type=""and"">
                  {conditionProject}
                  {conditionUnit}
                  <condition attribute=""bsd_douutien"" operator=""not-null"" />
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{enQueue.Id}"" />
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
                <order attribute=""bsd_dateorder"" descending=""false"" />
              </entity>
            </fetch>";
            EntityCollection enOpportunities = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            int priority = 1;
            foreach (var item in enOpportunities.Entities)
            {
                Entity enQueue_NewProority = new Entity(item.LogicalName, item.Id);
                if ((int)enQueue["bsd_douutien"] == 1 && priority == 1)
                {
                    enQueue_NewProority["bsd_douutien"] = priority;
                    enQueue_NewProority["statuscode"] = new OptionSetValue(100000004); // Queuing
                    this._service.Update(enQueue_NewProority);
                }
                else if ((int)enQueue["bsd_douutien"] == 1 && priority != 1)
                {
                    enQueue_NewProority["bsd_douutien"] = priority;
                    this._service.Update(enQueue_NewProority);
                }
                else if ((int)enQueue["bsd_douutien"] != 1)
                {
                    var priorityCurrent = (int)enQueue["bsd_douutien"];
                    if (priorityCurrent < (int)item["bsd_douutien"])
                    {
                        int flag = priorityCurrent + 1;
                        priority = flag == (int)item["bsd_douutien"] ? priorityCurrent : priority;
                        enQueue_NewProority["bsd_douutien"] = priority;
                        this._service.Update(enQueue_NewProority);
                    }
                }

                priority++;
            }
        }
        private void create_Refund(Entity enQueue)
        {
            Entity creRefund = new Entity("bsd_refund");
            string nameUnit = (string)enQueue["bsd_name"];
            //if (enQueue.Contains("bsd_unit"))
            //{
            //    Entity enUnit = this._service.Retrieve(((EntityReference)enQueue["bsd_unit"]).LogicalName, ((EntityReference)enQueue["bsd_unit"]).Id, new ColumnSet("bsd_name"));
            //    if (enUnit.Contains("bsd_name")) nameUnit = (string)enUnit["bsd_name"];
            //}
            creRefund["bsd_name"] = "Booking Refund - " + nameUnit;
            creRefund["bsd_customer"] = enQueue.Contains("bsd_customerid") ? (EntityReference)enQueue["bsd_customerid"] : null;
            creRefund["bsd_project"] = enQueue.Contains("bsd_project") ? (EntityReference)enQueue["bsd_project"] : null;
            creRefund["bsd_refundtype"] = new OptionSetValue(100000003);
            creRefund["bsd_unitno"] = enQueue.Contains("bsd_unit") ? (EntityReference)enQueue["bsd_unit"] : null;
            creRefund["bsd_queue"] = enQueue.ToEntityReference();
            decimal bsd_queuingfeepaid = ((Money)enQueue["bsd_queuingfeepaid"]).Value;
            creRefund["bsd_totalamountpaid"] = new Money(bsd_queuingfeepaid);
            creRefund["bsd_refundableamount"] = new Money(bsd_queuingfeepaid);
            _service.Create(creRefund);
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
            //var utcTime = utcTime.ToString("MM/dd/yyyy HH:mm:ss");
            //var localDateOnly = response.LocalTime.ToString("dd-MM-yyyy");
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
