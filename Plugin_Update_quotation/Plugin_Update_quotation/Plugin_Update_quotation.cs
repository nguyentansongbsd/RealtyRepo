using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_Update_quotation
{
    public class Plugin_Update_quotation : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity)) return;

            Entity target = context.InputParameters["Target"] as Entity;

            // 1. Kiểm tra Depth để tránh lặp vô tận (Infinite Loop)
            if (context.Depth > 1) return;

            // 2. Lấy data đầy đủ của Quote
            Entity quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
            Entity up_quote = new Entity(quote.LogicalName, quote.Id);

            decimal detailAmount = 0;
            Entity enUnit = null;

            // 3. Xử lý lấy giá từ Price Level
            if (quote.Contains("bsd_unitno") && quote.Contains("bsd_pricelevel"))
            {
                enUnit = service.Retrieve(((EntityReference)quote["bsd_unitno"]).LogicalName, ((EntityReference)quote["bsd_unitno"]).Id, new ColumnSet(true));
                EntityReference priceLevelRef = (EntityReference)quote["bsd_pricelevel"];

                var fetchXml = $@"
                <fetch distinct=""true"">
                  <entity name=""bsd_productpricelevel"">
                    <attribute name=""bsd_price"" alias=""prilist_price"" />
                    <filter>
                      <condition attribute=""bsd_product"" operator=""eq"" value=""{enUnit.Id}"" />
                      <condition attribute=""bsd_pricelevel"" operator=""eq"" value=""{priceLevelRef.Id}"" />
                    </filter>
                  </entity>
                </fetch>";

                EntityCollection rs_price = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs_price.Entities.Count > 0 && rs_price.Entities[0].Contains("prilist_price"))
                {
                    var aliased_money = (AliasedValue)rs_price.Entities[0]["prilist_price"];
                    Money moneyValue = (Money)aliased_money.Value;
                    detailAmount = moneyValue.Value;
                    up_quote["bsd_detailamount"] = moneyValue;
                }
            }

            // 4. Lấy các thông số cơ bản
            decimal discountAmount = quote.Contains("bsd_discountamount") ? ((Money)quote["bsd_discountamount"]).Value : 0;
            // Lưu ý: bsd_maintenancefeespercent thường là kiểu Decimal hoặc Double trong CRM
            decimal maintPercent = 0;
            if (quote.Contains("bsd_maintenancefeespercent"))
            {
                var val = quote["bsd_maintenancefeespercent"];
                maintPercent = (val is Money) ? ((Money)val).Value : (val is decimal ? (decimal)val : Convert.ToDecimal(val));
            }

            // 5. Tính toán theo Điều kiện bàn giao
            if (quote.Contains("bsd_handovercondition") && enUnit != null)
            {
                Entity handover = service.Retrieve(((EntityReference)quote["bsd_handovercondition"]).LogicalName, ((EntityReference)quote["bsd_handovercondition"]).Id, new ColumnSet(true));
                int bsd_method = handover.Contains("bsd_method") ? ((OptionSetValue)handover["bsd_method"]).Value : 0;

                decimal packageSellingAmount = 0;

                if (bsd_method == 100000001) // Fix Amount
                {
                    packageSellingAmount = handover.Contains("bsd_amount") ? ((Money)handover["bsd_amount"]).Value : 0;
                }
                else if (bsd_method == 100000002) // Percent
                {
                    decimal bsd_percent = handover.Contains("bsd_percent") ? (decimal)handover["bsd_percent"] : 0;
                    packageSellingAmount = (bsd_percent / 100m) * (detailAmount - discountAmount);
                }

                // Công thức tính toán chung
                decimal netAmount = detailAmount + packageSellingAmount - discountAmount;

                // Lấy thông tin từ Unit (Sản phẩm)
                decimal landValueUnit = enUnit.Contains("bsd_landvalueofunit") ? ((Money)enUnit["bsd_landvalueofunit"]).Value : 0;
                decimal netArea = enUnit.Contains("bsd_netsaleablearea") ? Convert.ToDecimal(enUnit["bsd_netsaleablearea"]) : 0;

                decimal landDeduction = landValueUnit * netArea;

                // Thuế VAT (10% trên số tiền sau khi trừ khấu trừ đất)
                decimal vat = (netAmount - landDeduction) * 0.1m;
                if (vat < 0) vat = 0;

                decimal maintenanceFee = netAmount * (maintPercent / 100m);

                // Gán giá trị vào Entity cập nhật
                up_quote["bsd_packagesellingamount"] = new Money(packageSellingAmount);
                up_quote["bsd_totalamountlessfreight"] = new Money(netAmount);
                up_quote["bsd_landvaluededuction"] = new Money(landDeduction);
                up_quote["bsd_vat"] = new Money(vat);
                up_quote["bsd_maintenancefees"] = new Money(maintenanceFee);
                up_quote["bsd_totalamountlessfreightaftervat"] = new Money(netAmount + vat);
                up_quote["bsd_totalamount"] = new Money(netAmount + vat + maintenanceFee);

                // 6. THỰC THI CẬP NHẬT
                service.Update(up_quote);
                trace.Trace("Đã cập nhật Quote ID: " + target.Id);
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
                ColumnSet = new ColumnSet("timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            });

            if (currentUserSettings.Entities.Count > 0)
                return (int?)currentUserSettings.Entities[0].Attributes["timezonecode"];
            return null;
        }
    }
}