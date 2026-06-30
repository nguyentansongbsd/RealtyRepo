using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_BreachManagement_Approve
{
    public class Plugin_BreachManagement_Approve : IPlugin
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
                Entity enBreachManagement = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_spa", "bsd_violatortype", "bsd_totalamount", "bsd_installment" }));
                int status = enBreachManagement.Contains("statuscode") ? ((OptionSetValue)enBreachManagement["statuscode"]).Value : -99;
                if (status == 100000000)  //Confirm
                {
                    CheckExistFile(enBreachManagement);
                }
                else if (status == 100000005)  //Complete
                {
                    EntityReference refSPA = (EntityReference)enBreachManagement["bsd_spa"];
                    Entity enSPA = service.Retrieve(refSPA.LogicalName, refSPA.Id, new ColumnSet(new string[] { "bsd_unitnumber", "bsd_customerid", "bsd_project", "bsd_totalamountpaid" }));
                    EntityReference refUnit = (EntityReference)enSPA["bsd_unitnumber"];

                    int bsd_violatortype = enBreachManagement.Contains("bsd_violatortype") ? ((OptionSetValue)enBreachManagement["bsd_violatortype"]).Value : -99;
                    if (bsd_violatortype == 100000000)   //Khách hàng
                    {
                        CreateMiscellaneous(enBreachManagement, refSPA, refUnit);

                        Entity upSPA = new Entity(refSPA.LogicalName, refSPA.Id);
                        upSPA["bsd_miscellaneous"] = enBreachManagement.Contains("bsd_totalamount") ? enBreachManagement["bsd_totalamount"] : null;
                        service.Update(upSPA);
                    }
                    else if (bsd_violatortype == 100000001)   //Chủ đầu tư
                    {
                        CreateRefund(enBreachManagement, enSPA, refUnit);
                    }
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private string GetURL()
        {

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_configgolive"">
                <attribute name=""bsd_configgoliveid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_url"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_url"" operator=""not-null"" />
                  <condition attribute=""bsd_name"" operator=""eq"" value=""BreachManagement_CheckFile"" />
                </filter>
              </entity>
            </fetch>";
            var rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                return (string)rs.Entities[0]["bsd_url"];
            }
            return null;
        }

        private void CheckExistFile(Entity enBreachManagement)
        {
            traceService.Trace("CheckExistFile");

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string flowUrl = GetURL();
                    if (string.IsNullOrWhiteSpace(flowUrl))
                        throw new InvalidPluginExecutionException("URL không hợp lệ, vui lồng kiểm tra lại cấu hình.");

                    var jsonData = new { recordId = enBreachManagement.Id.ToString() };

                    string jsonContent = JsonConvert.SerializeObject(jsonData);
                    HttpContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = client.PostAsync(flowUrl, content).GetAwaiter().GetResult();
                    string respContent = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode)
                    {
                        traceService.Trace($"PA trả lỗi: {response.StatusCode} || {respContent}");
                    }
                    else
                    {
                        dynamic result = JsonConvert.DeserializeObject(respContent);

                        bool hasPdf = result.hasPdf != null && Convert.ToBoolean(result.hasPdf);
                        string fileName = result.fileName != null ? (string)result.fileName : null;

                        traceService.Trace($"Has PDF: {hasPdf} || {fileName}");

                        if (!hasPdf)
                            throw new InvalidPluginExecutionException("Proposal Document has not been uploaded, please check again.");

                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        private void CreateMiscellaneous(Entity enBM, EntityReference refSPA, EntityReference refUnit)
        {
            traceService.Trace("CreateMiscellaneous");

            Entity newMisc = new Entity("bsd_miscellaneous");
            newMisc["bsd_name"] = $"Breach Management Miscellaneous - {refUnit.Name}";
            newMisc["bsd_spa"] = refSPA;
            newMisc["bsd_breachmanagement"] = enBM.ToEntityReference();
            newMisc["bsd_installment"] = enBM.Contains("bsd_installment") ? enBM["bsd_installment"] : null;
            newMisc["bsd_amount"] = enBM.Contains("bsd_totalamount") ? enBM["bsd_totalamount"] : null;

            newMisc.Id = Guid.NewGuid();
            service.Create(newMisc);
        }

        private void CreateRefund(Entity enBM, Entity enSPA, EntityReference refUnit)
        {
            traceService.Trace("CreateRefund");

            Entity newRefund = new Entity("bsd_refund");
            newRefund["bsd_name"] = $"Breach Management Refund-{refUnit.Name}";
            newRefund["bsd_optionentry"] = enSPA.ToEntityReference();
            newRefund["bsd_breachmanagement"] = enBM.ToEntityReference();
            newRefund["bsd_customer"] = enSPA.Contains("bsd_customerid") ? enSPA["bsd_customerid"] : null;
            newRefund["bsd_project"] = enSPA.Contains("bsd_project") ? enSPA["bsd_project"] : null;
            newRefund["bsd_refundtype"] = new OptionSetValue(100000004);    //Breach Refund
            newRefund["bsd_unitno"] = refUnit;
            newRefund["bsd_paymentactualtime"] = DateTime.UtcNow;
            newRefund["bsd_totalamountpaid"] = enSPA.Contains("bsd_totalamountpaid") ? enSPA["bsd_totalamountpaid"] : null;
            newRefund["bsd_refundableamount"] = enSPA.Contains("bsd_totalamountpaid") ? enSPA["bsd_totalamountpaid"] : null;
            newRefund["bsd_source"] = new OptionSetValue(100000002);

            newRefund.Id = Guid.NewGuid();
            service.Create(newRefund);
        }
    }
}