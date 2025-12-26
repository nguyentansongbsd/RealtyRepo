using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Queue_UpdatePriority
{
    public class Plugin_Queue_UpdatePriority : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationService _service;
        private IOrganizationServiceFactory _serviceFactory;
        private ITracingService _tracingService;
        public void Execute(IServiceProvider serviceProvider)
        {
            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (this._context.Depth > 3) return;

            if (this._context.InputParameters.Contains("Target") && this._context.InputParameters["Target"] is Entity)
            {
                Entity target = (Entity)this._context.InputParameters["Target"];
                Entity entityQueue = this._service.Retrieve(target.LogicalName, target.Id, new ColumnSet("bsd_collectedqueuingfee", "bsd_project", "bsd_unit", "bsd_queueforproject"));
                if ((bool)entityQueue["bsd_collectedqueuingfee"] == false) return;

                try
                {
                    _tracingService.Trace("Start Plugin_Queue_UpdatePriority");
                    int stt = 0;
                    int sut = 0;
                    int dut = 0;
                    getPriority(entityQueue, entityQueue.Id, ref stt, ref sut, ref dut);
                    bool isHadQueueing = checkStsQueue(entityQueue, entityQueue.Id);
                    this._tracingService.Trace("Co queueing ? " + isHadQueueing);
                    Entity queueItem = new Entity(entityQueue.LogicalName);
                    queueItem.Id = entityQueue.Id;
                    queueItem.Attributes["bsd_douutien"] = dut + 1;
                    if(entityQueue.Contains("bsd_queueforproject") && (bool)entityQueue["bsd_queueforproject"] == true)
                        queueItem.Attributes["bsd_sothutu"] =  stt + 1 ;
                    if(!entityQueue.Contains("bsd_queueforproject") || (entityQueue.Contains("bsd_queueforproject") && (bool)entityQueue["bsd_queueforproject"] == false))
                        queueItem.Attributes["bsd_souutien"] = sut + 1;
                    queueItem.Attributes["statuscode"] = isHadQueueing == false ? new OptionSetValue(100000004) : new OptionSetValue(100000003);//100000004: sts queueing; 100000003: sts waiting in queue
                    this._service.Update(queueItem);
                }
                catch (Exception ex)
                {
                    this._tracingService.Trace("Plugin_Queue_UpdatePriority: {0}", ex.ToString());
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        private void getPriority(Entity enQueue, Guid queueId, ref int stt, ref int sut, ref int dut)
        {
            string conditionUnit = enQueue.Contains("bsd_unit") ? $@"<condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_unit"]).Id}"" />" : $@"<condition attribute=""bsd_unit"" operator=""null"" />";
            string conditionProject = enQueue.Contains("bsd_project") ? $@"<condition attribute=""bsd_project"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_project"]).Id}"" />" : "";
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch aggregate=""true"">
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_douutien"" alias=""douutien"" aggregate=""max"" />
                <attribute name=""bsd_sothutu"" alias=""sothutu"" aggregate=""max"" />
                <attribute name=""bsd_souutien"" alias=""souutien"" aggregate=""max"" />
                <filter>
                  {conditionUnit}
                  {conditionProject}
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{queueId}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if(result.Entities.Count > 0)
            {
                if (result.Entities[0].Attributes.Contains("sothutu"))
                {
                    stt = ((AliasedValue)result.Entities[0]["sothutu"]).Value != null ? (int)((AliasedValue)result.Entities[0]["sothutu"]).Value : 0;
                }
                if (result.Entities[0].Attributes.Contains("souutien"))
                {
                    sut = ((AliasedValue)result.Entities[0]["souutien"]).Value != null ? (int)((AliasedValue)result.Entities[0]["souutien"]).Value : 0;
                }
                if (result.Entities[0].Attributes.Contains("douutien"))
                {
                    dut = ((AliasedValue)result.Entities[0]["douutien"]).Value != null ? (int)((AliasedValue)result.Entities[0]["douutien"]).Value : 0;
                }
            }
        }
        private bool checkStsQueue(Entity enQueue, Guid queueId)
        {
            // sts queueing = 100000004
            string conditionUnit = enQueue.Contains("bsd_unit") ? $@"<condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_unit"]).Id}"" />" : $@"<condition attribute=""bsd_unit"" operator=""null"" />";
            string conditionProject = enQueue.Contains("bsd_project") ? $@"<condition attribute=""bsd_project"" operator=""eq"" value=""{((EntityReference)enQueue["bsd_project"]).Id}"" />" : "";
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <filter>
                  {conditionUnit}
                  {conditionProject}
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{queueId}"" />
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000004"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            this._tracingService.Trace("fetch check sts queue: " + fetchXml);
            if(result.Entities.Count > 0)
                return true;
            else
                return false;
        }
    }
}
