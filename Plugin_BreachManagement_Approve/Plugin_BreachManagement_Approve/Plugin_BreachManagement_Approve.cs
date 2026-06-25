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
                Entity enBreachManagement = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_spa", "bsd_violatortype", "bsd_totalamount" }));
                int status = enBreachManagement.Contains("statuscode") ? ((OptionSetValue)enBreachManagement["statuscode"]).Value : -99;
                if (status == 100000000)  //Confirm
                {
                    CheckExistFile(enBreachManagement);
                }
                else if (status == 100000005)  //Complete
                {
                    int bsd_violatortype = enBreachManagement.Contains("bsd_violatortype") ? ((OptionSetValue)enBreachManagement["bsd_violatortype"]).Value : -99;
                    if (bsd_violatortype == 100000000)   //Khách hàng
                    {
                        EntityReference SPA = (EntityReference)enBreachManagement["bsd_spa"];
                        Entity upSPA = new Entity(SPA.LogicalName, SPA.Id);
                        upSPA["bsd_miscellaneous"] = enBreachManagement.Contains("bsd_totalamount") ? enBreachManagement["bsd_totalamount"] : null;
                        service.Update(upSPA);
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
    }
}