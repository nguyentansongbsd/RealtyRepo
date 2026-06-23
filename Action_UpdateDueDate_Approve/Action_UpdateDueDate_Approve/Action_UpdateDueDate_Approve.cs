using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_UpdateDueDate_Approve
{
    public class Action_UpdateDueDate_Approve : IPlugin
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

                EntityReference refUpdateDueDate = (EntityReference)context.InputParameters["Target"];

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_updateduedatedetail"">
                    <attribute name=""bsd_updateduedatedetailid"" />
                    <attribute name=""bsd_ra"" />
                    <attribute name=""bsd_spa"" />
                    <attribute name=""bsd_installment"" />
                    <attribute name=""bsd_duedateold"" />
                    <attribute name=""bsd_duedatenew"" />
                    <filter>
                      <condition attribute=""bsd_updateduedate"" operator=""eq"" value=""{refUpdateDueDate.Id}"" />
                    </filter>
                    <link-entity name=""bsd_paymentschemedetail"" from=""bsd_paymentschemedetailid"" to=""bsd_installment"" alias=""ins"">
                      <attribute name=""statuscode"" />
                      <attribute name=""bsd_ordernumber"" />
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    //  trùng đợt 
                    var duplicateInstallments = rs.Entities
                                                .Where(e => e.Contains("bsd_installment") && e["bsd_installment"] != null)
                                                .GroupBy(e => ((EntityReference)e["bsd_installment"]).Id)
                                                .Where(g => g.Count() > 1);
                    if (duplicateInstallments.Any())
                        throw new InvalidPluginExecutionException("Trùng đợt, vui lòng kiểm tra lại.");

                    Dictionary<Guid, DateTime> newDueDateMap = rs.Entities
                    .Where(x => x.Contains("bsd_installment") && x.Contains("bsd_duedatenew"))
                    .ToDictionary(
                        x => ((EntityReference)x["bsd_installment"]).Id,
                        x => RetrieveLocalTimeFromUTCTime((DateTime)x["bsd_duedatenew"], service)
                    );

                    foreach (Entity enDetail in rs.Entities)
                    {
                        string logicalName = string.Empty;
                        Entity enContract = null;
                        if (enDetail.Contains("bsd_ra"))
                        {
                            logicalName = "bsd_reservationcontract";
                            EntityReference refRA = (EntityReference)enDetail["bsd_ra"];
                            enContract = service.Retrieve(refRA.LogicalName, refRA.Id, new ColumnSet(new string[] { "bsd_name", "statuscode" }));
                        }
                        else
                        {
                            logicalName = "bsd_optionentry";
                            EntityReference refSPA = (EntityReference)enDetail["bsd_spa"];
                            enContract = service.Retrieve(refSPA.LogicalName, refSPA.Id, new ColumnSet(new string[] { "bsd_name", "statuscode" }));
                        }

                        CheckContract(enDetail, enContract, logicalName, newDueDateMap);
                    }

                    foreach (var enDetail in rs.Entities)
                    {
                        Entity upDetail = new Entity(enDetail.LogicalName, enDetail.Id);
                        upDetail["statuscode"] = new OptionSetValue(100000000); //Approved
                        service.Update(upDetail);

                        UpdateNewDueDate(enDetail, newDueDateMap);
                    }
                }
                else
                {
                    //  1. Update Duedate master phải tồn tại update Duedate detail
                    throw new InvalidPluginExecutionException("Không có chi tiết cập nhật ngày đến hạn. Vui lòng kiểm tra lại thông tin.");
                }

                Entity upUpdateDueDate = new Entity(refUpdateDueDate.LogicalName, refUpdateDueDate.Id);
                upUpdateDueDate["statuscode"] = new OptionSetValue(100000000); //Approved
                service.Update(upUpdateDueDate);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }


        private void CheckContract(Entity enDetail, Entity enContract, string logicalName, Dictionary<Guid, DateTime> newDueDateMap)
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

            string contractName = enContract.Contains("bsd_name") ? (string)enContract["bsd_name"] : string.Empty;

            // check đợt paid
            if (enDetail.Contains("ins.statuscode") && enDetail["ins.statuscode"] != null)
            {
                int statusCode = ((OptionSetValue)((AliasedValue)enDetail["ins.statuscode"]).Value).Value;
                if (statusCode == 100000001)    //Paid
                {
                    int? bsd_ordernumber = enDetail.Contains("ins.bsd_ordernumber") ? (int?)((AliasedValue)enDetail["ins.bsd_ordernumber"]).Value : null;
                    throw new InvalidPluginExecutionException($"Mã hợp đồng '{contractName}' có Đợt '{bsd_ordernumber}' đã thanh toán hoàn tất. Không thể cập nhật ngày đến hạn.");
                }
            }

            //  4. DueDate Đợt n mới > DueDate đợt n cũ
            if (enDetail.Contains("bsd_duedateold") && enDetail.Contains("bsd_duedatenew") &&
                ((DateTime)enDetail["bsd_duedateold"]).Date >= ((DateTime)enDetail["bsd_duedatenew"]).Date)
                throw new InvalidPluginExecutionException("Ngày đến hạn mới không được nhỏ hơn hoặc bằng ngày đến hạn cũ.");

            //  2. Kiểm tra ngày đến hạn đợt n có lớn hơn đợt n -1
            //  3. Kiểm tra ngày đến hạn đợt n có nhỏ hơn đợt n+1
            CheckInsDueDate(enContract, contractName, logicalName, newDueDateMap);
        }

        private void CheckInsDueDate(Entity enContract, string contractName, string logicalName, Dictionary<Guid, DateTime> newDueDateMap)
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
                        DueDate = newDueDateMap.ContainsKey(x.Id) ? newDueDateMap[x.Id] : RetrieveLocalTimeFromUTCTime((DateTime)x["bsd_duedate"], service)
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

        private void UpdateNewDueDate(Entity enDetail, Dictionary<Guid, DateTime> newDueDateMap)
        {
            traceService.Trace("UpdateNewDueDate");

            EntityReference refIns = (EntityReference)enDetail["bsd_installment"];

            if (refIns == null || !newDueDateMap.ContainsKey(refIns.Id))
                return;

            DateTime newDate = newDueDateMap[refIns.Id];

            Entity updateIns = new Entity(refIns.LogicalName, refIns.Id);
            updateIns["bsd_duedate"] = newDate;
            service.Update(updateIns);
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