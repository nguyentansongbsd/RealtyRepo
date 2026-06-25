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
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        ITracingService traceService = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");

                string step = context.InputParameters.Contains("step") && !string.IsNullOrEmpty((string)context.InputParameters["step"]) ?
                            (string)context.InputParameters["step"] : string.Empty;
                Guid userId = context.InputParameters.Contains("userId") && !string.IsNullOrEmpty((string)context.InputParameters["userId"]) ?
                                Guid.Parse((string)context.InputParameters["userId"]) : context.UserId;
                traceService.Trace($"userId: {userId} || {step}");
                service = factory.CreateOrganizationService(userId);

                EntityReference refUpdateDueDate = (EntityReference)context.InputParameters["Target"];
                switch (step)
                {
                    case "b1":
                        Step_B1(refUpdateDueDate);
                        break;

                    case "b2":
                        Step_B2();
                        break;

                    case "error":
                        Step_Error();
                        break;

                    default:
                        Entity upUpdateDueDate2 = new Entity(refUpdateDueDate.LogicalName, refUpdateDueDate.Id);
                        upUpdateDueDate2["statuscode"] = new OptionSetValue(100000000); //Approved
                        upUpdateDueDate2["bsd_powerautomate"] = false;
                        upUpdateDueDate2["bsd_paprocess"] = null;
                        service.Update(upUpdateDueDate2);
                        break;
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void Step_B1(EntityReference refUpdateDueDate)
        {
            traceService.Trace("Step_B1");

            string paProcess = context.InputParameters.Contains("paProcess") && !string.IsNullOrEmpty((string)context.InputParameters["paProcess"]) ?
                (string)context.InputParameters["paProcess"] : string.Empty;

            Entity enUpdateDueDate = service.Retrieve(refUpdateDueDate.LogicalName, refUpdateDueDate.Id, new ColumnSet(new string[] { "bsd_powerautomate", "bsd_paprocess" }));
            bool isPA = enUpdateDueDate.Contains("bsd_powerautomate") ? (bool)enUpdateDueDate["bsd_powerautomate"] : false;
            if (isPA && enUpdateDueDate.Contains("bsd_paprocess") && (string)enUpdateDueDate["bsd_paprocess"] != paProcess)
                throw new InvalidPluginExecutionException("Record này đang được thực hiện ở tiến trình khác. Vui lòng kiểm tra lại.");

            Entity upUpdateDueDate = new Entity(refUpdateDueDate.LogicalName, refUpdateDueDate.Id);
            upUpdateDueDate["bsd_powerautomate"] = true;
            upUpdateDueDate["bsd_paprocess"] = paProcess;
            service.Update(upUpdateDueDate);
        }

        private void Step_B2()
        {
            traceService.Trace("Step_B2");

            string updateDueDateDetailId = context.InputParameters.Contains("updateDueDateDetailId") && !string.IsNullOrEmpty((string)context.InputParameters["updateDueDateDetailId"]) ?
                                (string)context.InputParameters["updateDueDateDetailId"] : string.Empty;
            Entity enDetail = service.Retrieve("bsd_updateduedatedetail", Guid.Parse(updateDueDateDetailId), new ColumnSet(new string[] { "bsd_ra", "bsd_spa",
                                                "bsd_installment", "bsd_duedateold", "bsd_duedatenew"}));

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

            EntityReference refIns = (EntityReference)enDetail["bsd_installment"];
            DateTime newDate = RetrieveLocalTimeFromUTCTime((DateTime)enDetail["bsd_duedatenew"], service);

            CheckContract(enDetail, enContract, logicalName, refIns, newDate);
            UpdateNewDueDate(refIns, newDate);

            Entity upDetail = new Entity(enDetail.LogicalName, enDetail.Id);
            upDetail["statuscode"] = new OptionSetValue(100000000); //Approved
            service.Update(upDetail);
        }

        private void Step_Error()
        {
            traceService.Trace("Step_Error");

            string detailId = context.InputParameters.Contains("updateDueDateDetailId") && !string.IsNullOrEmpty((string)context.InputParameters["updateDueDateDetailId"]) ?
                    (string)context.InputParameters["updateDueDateDetailId"] : string.Empty;
            string error = context.InputParameters.Contains("error") && !string.IsNullOrEmpty((string)context.InputParameters["error"]) ?
                    (string)context.InputParameters["error"] : string.Empty;

            Entity upDetailError = new Entity("bsd_updateduedatedetail", Guid.Parse(detailId));
            upDetailError["bsd_error"] = error;
            service.Update(upDetailError);
        }


        private void CheckContract(Entity enDetail, Entity enContract, string logicalName, EntityReference refIns, DateTime newDate)
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
            Entity enIns = service.Retrieve(refIns.LogicalName, refIns.Id, new ColumnSet(new string[] { "statuscode", "bsd_ordernumber" }));
            int statusCode = ((OptionSetValue)enIns["statuscode"]).Value;
            if (statusCode == 100000001)    //Paid
            {
                int? bsd_ordernumber = enIns.Contains("bsd_ordernumber") ? (int?)enIns["bsd_ordernumber"] : null;
                throw new InvalidPluginExecutionException($"Mã hợp đồng '{contractName}' có Đợt '{bsd_ordernumber}' đã thanh toán hoàn tất. Không thể cập nhật ngày đến hạn.");
            }

            //  4. DueDate Đợt n mới > DueDate đợt n cũ
            if (enDetail.Contains("bsd_duedateold") && enDetail.Contains("bsd_duedatenew") &&
                ((DateTime)enDetail["bsd_duedateold"]).Date >= ((DateTime)enDetail["bsd_duedatenew"]).Date)
                throw new InvalidPluginExecutionException("Ngày đến hạn mới không được nhỏ hơn hoặc bằng ngày đến hạn cũ.");

            //  2. Kiểm tra ngày đến hạn đợt n có lớn hơn đợt n -1
            //  3. Kiểm tra ngày đến hạn đợt n có nhỏ hơn đợt n+1
            CheckInsDueDate(enContract, contractName, logicalName, refIns, newDate);
        }

        private void CheckInsDueDate(Entity enContract, string contractName, string logicalName, EntityReference refIns, DateTime newDate)
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
                        DueDate = refIns.Id == x.Id ? newDate : RetrieveLocalTimeFromUTCTime((DateTime)x["bsd_duedate"], service)
                    })
                    .OrderBy(x => x.Order)
                    .ToList();

                for (int i = 0; i < rs.Entities.Count - 1; i++)
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

        private void UpdateNewDueDate(EntityReference refIns, DateTime newDate)
        {
            traceService.Trace("UpdateNewDueDate");

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