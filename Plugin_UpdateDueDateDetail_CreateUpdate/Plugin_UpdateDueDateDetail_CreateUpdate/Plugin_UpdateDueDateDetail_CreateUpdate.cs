using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_UpdateDueDateDetail_CreateUpdate
{
    public class Plugin_UpdateDueDateDetail_CreateUpdate : IPlugin
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
                Entity enDetail = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_ra", "bsd_spa", "bsd_installmentnumber", "bsd_installment",
                "bsd_updateduedate", "bsd_duedateold", "bsd_duedatenew"}));

                string logicalName = string.Empty;
                Entity enContract = null;
                Entity upDetail = new Entity(enDetail.LogicalName, enDetail.Id);
                if (enDetail.Contains("bsd_ra"))
                {
                    logicalName = "bsd_reservationcontract";
                    EntityReference refRA = (EntityReference)enDetail["bsd_ra"];
                    enContract = service.Retrieve(refRA.LogicalName, refRA.Id, new ColumnSet(new string[] { "statuscode", "bsd_unitno", "bsd_name" }));
                    upDetail["bsd_units"] = enContract.Contains("bsd_unitno") ? enContract["bsd_unitno"] : null;
                }
                else
                {
                    logicalName = "bsd_optionentry";
                    EntityReference refSPA = (EntityReference)enDetail["bsd_spa"];
                    enContract = service.Retrieve(refSPA.LogicalName, refSPA.Id, new ColumnSet(new string[] { "statuscode", "bsd_unitnumber", "bsd_name" }));
                    upDetail["bsd_units"] = enContract.Contains("bsd_unitnumber") ? enContract["bsd_unitnumber"] : null;
                }

                if ("Create".Equals(context.MessageName))
                {
                    ForCreate(enDetail, enContract, logicalName, ref upDetail);
                }
                else
                {
                    ForUpdate(enDetail, enContract, logicalName, ref upDetail);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void ForCreate(Entity enDetail, Entity enContract, string logicalName, ref Entity upDetail)
        {
            traceService.Trace("ForCreate");

            if (enDetail.Contains("bsd_installmentnumber") && !enDetail.Contains("bsd_installment"))
            {
                int bsd_installmentnumber = (int)enDetail["bsd_installmentnumber"];

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch>
                              <entity name=""bsd_paymentschemedetail"">
                                <attribute name=""bsd_paymentschemedetailid"" />
                                <attribute name=""bsd_name"" />
                                <attribute name=""bsd_duedate"" />
                                <filter>
                                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                  <condition attribute=""{logicalName}"" operator=""eq"" value=""{enContract.Id}"" />
                                  <condition attribute=""bsd_ordernumber"" operator=""eq"" value=""{bsd_installmentnumber}"" />
                                </filter>
                              </entity>
                            </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    upDetail["bsd_installment"] = rs[0].ToEntityReference();
                    upDetail["bsd_duedateold"] = rs[0].Contains("bsd_duedate") ? rs[0]["bsd_duedate"] : null;
                    enDetail["bsd_installment"] = upDetail["bsd_installment"];
                    enDetail["bsd_duedateold"] = upDetail["bsd_duedateold"];
                }
            }
            service.Update(upDetail);

            CheckContract(enDetail, enContract, logicalName);
        }

        private void ForUpdate(Entity enDetail, Entity enContract, string logicalName, ref Entity upDetail)
        {
            traceService.Trace("ForUpdate");

            if (!enDetail.Contains("bsd_installmentnumber") && enDetail.Contains("bsd_installment"))
            {
                EntityReference refIns = (EntityReference)enDetail["bsd_installment"];
                Entity enIns = service.Retrieve(refIns.LogicalName, refIns.Id, new ColumnSet(new string[] { "bsd_ordernumber" }));
                upDetail["bsd_installmentnumber"] = enIns.Contains("bsd_ordernumber") ? enIns["bsd_ordernumber"] : null;
            }
            service.Update(upDetail);

            CheckContract(enDetail, enContract, logicalName);
        }

        private void CheckContract(Entity enDetail, Entity enContract, string logicalName)
        {
            traceService.Trace("CheckContract");

            //1. Chỉ import thành công các HĐ chưa thanh lý
            int status = enContract.Contains("statuscode") ? ((OptionSetValue)enContract["statuscode"]).Value : -99;
            if ("bsd_reservationcontract".Equals(enContract.LogicalName))
            {
                if (status == 100000004) //Terminedted
                    throw new InvalidPluginExecutionException("Hợp đồng đã thanh lý. Không thể thực hiện thao tác.");
            }
            else
            {
                if (status == 100000014) //Terminated
                    throw new InvalidPluginExecutionException("Hợp đồng đã thanh lý. Không thể thực hiện thao tác.");
            }

            EntityReference refMaster = (EntityReference)enDetail["bsd_updateduedate"];
            Entity enMaster = service.Retrieve(refMaster.LogicalName, refMaster.Id, new ColumnSet(new string[] { "statuscode" }));

            //6. Import mà Master đã duyệt
            int statusMaster = enMaster.Contains("statuscode") ? ((OptionSetValue)enMaster["statuscode"]).Value : -99;
            if (statusMaster == 100000000) //Approved
                throw new InvalidPluginExecutionException("Update Due Date đã được duyệt.");

            EntityReference refIns = (EntityReference)enDetail["bsd_installment"];

            //7. Import trùng đợt 
            CheckDuplicateIns(enDetail, enMaster, refIns);

            //5. Đợt update không phải là đợt bàn giao/đợt cuối
            Entity enIns = service.Retrieve(refIns.LogicalName, refIns.Id, new ColumnSet(new string[] { "bsd_pinkbookhandover", "bsd_lastinstallment", "bsd_ordernumber",
                "bsd_duedate", "statuscode" }));
            string contractName = enContract.Contains("bsd_name") ? (string)enContract["bsd_name"] : string.Empty;

            // check đợt paid
            int statusIns = enIns.Contains("statuscode") ? ((OptionSetValue)enIns["statuscode"]).Value : -99;
            if (statusIns == 100000001) //Paid
            {
                int? bsd_ordernumber = enIns.Contains("bsd_ordernumber") ? (int?)enIns["bsd_ordernumber"] : null;
                throw new InvalidPluginExecutionException($"Mã hợp đồng '{contractName}' có Đợt '{bsd_ordernumber}' đã thanh toán hoàn tất. Không thể cập nhật ngày đến hạn.");
            }

            if (enIns.Contains("bsd_pinkbookhandover") && (bool)enIns["bsd_pinkbookhandover"])
                throw new InvalidPluginExecutionException("Đợt cập nhật đang thuộc đợt bàn giao. Không thể thực hiện cập nhật");

            if (enIns.Contains("bsd_lastinstallment") && (bool)enIns["bsd_lastinstallment"])
                throw new InvalidPluginExecutionException("Đợt cập nhật đang thuộc đợt cuối. Không thể thực hiện cập nhật");

            //4. DueDate Đợt n mới <= DueDate đợt n cũ
            if (enDetail.Contains("bsd_duedateold") && enDetail.Contains("bsd_duedatenew") &&
                ((DateTime)enDetail["bsd_duedateold"]).Date >= ((DateTime)enDetail["bsd_duedatenew"]).Date)
                throw new InvalidPluginExecutionException("Ngày đến hạn mới không được nhỏ hơn hoặc bằng ngày đến hạn cũ.");

            //  2. Ngày đến hạn đợt nhỏ hơn đợt trước nó
            //  3.Ngày đến hạn đợt lớn hơn đợt sau nó
            CheckInsDueDate(enDetail, contractName, enContract, logicalName, refIns);
        }

        private void CheckDuplicateIns(Entity enDetail, Entity enMaster, EntityReference refIns)
        {
            traceService.Trace("CheckDuplicateIns");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_updateduedatedetail"">
                <attribute name=""bsd_updateduedatedetailid"" />
                <filter>
                  <condition attribute=""bsd_updateduedate"" operator=""eq"" value=""{enMaster.Id}"" />
                  <condition attribute=""bsd_installment"" operator=""eq"" value=""{refIns.Id}"" />
                  <condition attribute=""bsd_updateduedatedetailid"" operator=""ne"" value=""{enDetail.Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                throw new InvalidPluginExecutionException("Trùng đợt, vui lòng kiểm tra lại.");
        }

        private void CheckInsDueDate(Entity enDetail, string contractName, Entity enContract, string logicalName, EntityReference refIns)
        {
            traceService.Trace("CheckInsDueDate");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_ordernumber"" />
                <attribute name=""bsd_duedate"" />
                <filter>
                  <condition attribute=""{logicalName}"" operator=""eq"" value=""{enContract.Id}""/>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                var newListIns = rs.Entities
                .Select(x => new
                {
                    Id = x.Id,
                    Order = x["bsd_ordernumber"],
                    Name = x["bsd_name"],
                    DueDate = refIns.Id == x.Id ? RetrieveLocalTimeFromUTCTime((DateTime)enDetail["bsd_duedatenew"], service) : RetrieveLocalTimeFromUTCTime((DateTime)x["bsd_duedate"], service)
                })
                .OrderBy(x => x.Order)
                .ToList();

                for (int i = 0; i < newListIns.Count - 1; i++)
                {
                    if (newListIns[i].DueDate.Date > newListIns[i + 1].DueDate.Date)
                    {
                        traceService.Trace("" + newListIns[i].DueDate.Date);

                        string currentName = (string)newListIns[i].Name;
                        string nextName = (string)newListIns[i + 1].Name;

                        throw new InvalidPluginExecutionException($"Mã hợp đồng '{contractName}' có ngày đến hạn của '{currentName}' đang lớn hơn '{nextName}'. Vui lòng kiểm tra thông tin.");
                    }
                }
            }
        }

        public DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService ser)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(ser);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            LocalTimeFromUtcTimeResponse response = (LocalTimeFromUtcTimeResponse)ser.Execute(request);
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