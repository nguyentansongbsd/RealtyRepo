using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_MatchUnit_Approve
{
    public class Action_MatchUnit_Approve : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private IOrganizationServiceFactory _serviceFactory;

        EntityReference targetEntityRef = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);
            try
            {
                this.targetEntityRef = (EntityReference)_context.InputParameters["Target"];
                Entity enMatchUnit = this._service.Retrieve(this.targetEntityRef.LogicalName, this.targetEntityRef.Id, new ColumnSet("bsd_queue",
                    "bsd_unit", "bsd_pricelist", "bsd_usableareasqm", "bsd_builtupareasqm", "bsd_usableunitprice", "bsd_builtupunitprice",
                    "bsd_price", "bsd_project"));

                DateTime dateApprove = DateTime.Now;
                UpdateMatchUnit(dateApprove);
                UpdateQueue(enMatchUnit, dateApprove);
                UpdatePriorityQueueProject(enMatchUnit);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        // Cập nhật thông tin Ráp căn
        private void UpdateMatchUnit(DateTime dateApprove)
        {
            this._tracingService.Trace("Start Update Match Unit");    
            Entity matchUnit = new Entity(this.targetEntityRef.LogicalName, this.targetEntityRef.Id);
            matchUnit["statuscode"] = new OptionSetValue(100000001); // Approved
            matchUnit["bsd_approver"] = new EntityReference("systemuser", this._context.UserId);
            matchUnit["bsd_approvedate"] = RetrieveLocalTimeFromUTCTime(dateApprove, this._service);
            this._service.Update(matchUnit);
            this._tracingService.Trace("End Update Match Unit");
        }
        // Cập nhật thông tin vào bản ghi Giữ chỗ
        private void UpdateQueue(Entity enMatchUnit, DateTime dateApprove)
        {
            this._tracingService.Trace("Start Update Queue");
            int sut = 0;
            int dut = 0;
            bool isHasStsQueuing = CheckStsQueuingWithUnit(enMatchUnit);
            getPriority(enMatchUnit, ref sut, ref dut);
            this._tracingService.Trace($"sut: {sut}, dut: {dut}");
            Entity queue = new Entity(((EntityReference)enMatchUnit["bsd_queue"]).LogicalName, ((EntityReference)enMatchUnit["bsd_queue"]).Id);
            queue["bsd_unit"] = enMatchUnit.Contains("bsd_unit") ? enMatchUnit["bsd_unit"] : null;
            queue["bsd_pricelist"] = enMatchUnit.Contains("bsd_pricelist") ? enMatchUnit["bsd_pricelist"] : null;
            queue["bsd_usableareasqm"] = enMatchUnit.Contains("bsd_usableareasqm") ? enMatchUnit["bsd_usableareasqm"] : null;
            queue["bsd_builtupareasqm"] = enMatchUnit.Contains("bsd_builtupareasqm") ? enMatchUnit["bsd_builtupareasqm"] : null;
            queue["bsd_usableunitprice"] = enMatchUnit.Contains("bsd_usableunitprice") ? enMatchUnit["bsd_usableunitprice"] : null;
            queue["bsd_builtupunitprice"] = enMatchUnit.Contains("bsd_builtupunitprice") ? enMatchUnit["bsd_builtupunitprice"] : null;
            queue["bsd_price"] = enMatchUnit.Contains("bsd_price") ? enMatchUnit["bsd_price"] : null;
            queue["bsd_dateorder"] = RetrieveLocalTimeFromUTCTime(dateApprove, this._service);
            queue["bsd_souutien"] = sut;
            queue["bsd_douutien"] = dut;
            queue["statuscode"] = isHasStsQueuing == true ? new OptionSetValue(100000003) : new OptionSetValue(100000007); // 100000003 - Wariting, 100000007 - Confirmed
            this._service.Update(queue);
            this._tracingService.Trace("End Update Queue");
        }
        // Lấy số ưu tiên và độ ưu tiên
        private void getPriority(Entity enMatchUnit, ref int sut, ref int dut)
        {
            if (!enMatchUnit.Contains("bsd_unit")) return;
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_douutien"" />
                <attribute name=""bsd_souutien"" />
                <attribute name=""statuscode"" />
                <filter>
                  <condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enMatchUnit["bsd_unit"]).Id}"" />
                  <condition attribute=""statuscode"" operator=""in"">
                      <value>100000003</value>
                      <value>100000004</value>
                      <value>100000005</value>
                  </condition>
                </filter>
              </entity>
            </fetch>";
            EntityCollection result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            this._tracingService.Trace("fetch: " + fetchXml);
            if (result.Entities.Count > 0) // Case da co Giu cho
            {
                int SUT_Max = 0;
                int DUT_Max = 0;
                SUT_Max = result.Entities
                    .Where(x => x.Contains("bsd_souutien") && x["bsd_souutien"] != null)
                    .Select(x => (int)x["bsd_souutien"])
                    .DefaultIfEmpty(0)
                    .Max();

                DUT_Max = result.Entities
                    .Where(x =>
                        x.Contains("bsd_douutien") &&
                        x["bsd_douutien"] != null &&
                        x.Contains("statuscode") &&
                        (
                            ((OptionSetValue)x["statuscode"]).Value == 100000003 ||
                            ((OptionSetValue)x["statuscode"]).Value == 100000004
                        )
                    )
                    .Select(x => (int)x["bsd_douutien"])
                    .DefaultIfEmpty(0)
                    .Max();
                this._tracingService.Trace($"SUT_Max: {SUT_Max}, DUT_Max: {DUT_Max}");
                sut = SUT_Max + 1;
                dut = DUT_Max + 1;
            }
            else // Case chua co Giu cho
            {
                sut = 1;
                dut = 1;
            }
        }
        // Cập nhật thông tin ưu tiên vào giữ chỗ dự án còn lại
        private void UpdatePriorityQueueProject(Entity enMatchUnit)
        {
            if (!enMatchUnit.Contains("bsd_project")) return;
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_douutien"" />
                <filter>
                  <condition attribute=""bsd_project"" operator=""eq"" value=""{((EntityReference)enMatchUnit["bsd_project"]).Id}"" />
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{((EntityReference)enMatchUnit["bsd_queue"]).Id}"" />                  
                  <condition attribute=""bsd_unit"" operator=""null"" />
                  <condition attribute=""bsd_queueforproject"" operator=""eq"" value=""1"" />
                </filter>
                
              </entity>
            </fetch>";
            //<filter>
            //      <condition attribute=""statuscode"" operator=""in"">
            //        <value>100000003</value>
            //      </condition>
            //    </filter>
            EntityCollection result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (result.Entities.Count > 0)
            {
                foreach (var item in result.Entities.OrderBy(x => x.GetAttributeValue<int>("bsd_douutien")))
                {
                    Entity queueProject = new Entity(item.LogicalName, item.Id);
                    //if ((int)item["bsd_douutien"] == 2)// Neu do uu tien = 2 thi cap nhat lai trang thai thanh queuing va giam do uu tien = 1. với các GC khác thì chỉ giảm độ ưu tiên và không thay đổi status
                    //    queueProject["statuscode"] = new OptionSetValue(100000004); // 100000004 - Queuing
                    queueProject["bsd_douutien"] = (int)item["bsd_douutien"] - 1;
                    this._service.Update(queueProject);
                }
            }
        }
        private bool CheckStsQueuingWithUnit(Entity enMatchUnit)
        {
            if (!enMatchUnit.Contains("bsd_unit")) return false;
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <filter>
                  <condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enMatchUnit["bsd_unit"]).Id}"" />
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000007"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (result.Entities.Count > 0)
                return true;
            return false;
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
