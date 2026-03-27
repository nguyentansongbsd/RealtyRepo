using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Units_CreateQuotation
{
    public class Action_Units_CreateQuotation : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 1) return;

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity enUnit = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_name", "bsd_projectcode",
                "bsd_pricelevel", "bsd_taxcode", "bsd_unittype", "bsd_netsaleablearea", "bsd_numberofmonthspaidmf", "bsd_managementamountmonth", "bsd_price",
                "bsd_landvalueofunit", "bsd_maintenancefeespercent"}));
                int status = enUnit.Contains("statuscode") ? ((OptionSetValue)enUnit["statuscode"]).Value : -99;
                if (status != 100000000) //Available
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_unit"));

                Guid idOE = CreateQuotation(enUnit, target);

                traceService.Trace("done");
                context.OutputParameters["id"] = idOE.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private object GetValidFieldValue(Entity enUnit, string field)
        {
            return enUnit.Contains(field) ? enUnit[field] : null;
        }

        private Guid CreateQuotation(Entity enUnit, EntityReference target)
        {
            traceService.Trace("CreateQuotation");

            Entity newQuotation = new Entity("bsd_quotation");
            newQuotation["bsd_name"] = $"Quotation {GetValidFieldValue(enUnit, "bsd_name")}";
            newQuotation["bsd_date"] = DateTime.UtcNow;
            newQuotation["bsd_project"] = GetValidFieldValue(enUnit, "bsd_projectcode");
            newQuotation["bsd_phaseslaunch"] = GetPhasesLaunch(enUnit, ref newQuotation);
            newQuotation["bsd_pricelevel"] = GetValidFieldValue(enUnit, "bsd_pricelevel");
            newQuotation["bsd_taxcode"] = GetValidFieldValue(enUnit, "bsd_taxcode");
            //newQuotation["bsd_unittype"] = GetValidFieldValue(enUnit, "bsd_unittype");
            newQuotation["bsd_unitnumber"] = target;
            decimal bsd_netsaleablearea = enUnit.Contains("bsd_netsaleablearea") ? (decimal)enUnit["bsd_netsaleablearea"] : 0;
            newQuotation["bsd_netusablearea"] = bsd_netsaleablearea;

            //Price
            newQuotation["bsd_detailamount"] = enUnit.Contains("bsd_price") ? enUnit["bsd_price"] : new Money(0);
            newQuotation["bsd_totalamountlessfreight"] = enUnit.Contains("bsd_price") ? enUnit["bsd_price"] : new Money(0);
            newQuotation["bsd_discountamount"] = new Money(0);

            decimal bsd_landvalueofunit = enUnit.Contains("bsd_landvalueofunit") ? ((Money)enUnit["bsd_landvalueofunit"]).Value : 0;
            newQuotation["bsd_landvaluededuction"] = new Money(bsd_landvalueofunit * bsd_netsaleablearea);

            //Management Fee Information
            #region Management Fee Information
            int bsd_numberofmonthspaidmf = 0;
            if (enUnit.Contains("bsd_numberofmonthspaidmf"))
            {
                bsd_numberofmonthspaidmf = (int)enUnit["bsd_numberofmonthspaidmf"];
                newQuotation["bsd_numberofmonthspaidmf"] = bsd_numberofmonthspaidmf;
            }
            decimal bsd_managementamountmonth = enUnit.Contains("bsd_managementamountmonth") ? ((Money)enUnit["bsd_managementamountmonth"]).Value : 0;
            newQuotation["bsd_managementfee"] = new Money(bsd_netsaleablearea * bsd_managementamountmonth * bsd_numberofmonthspaidmf);
            newQuotation["bsd_maintenancefeespercent"] = GetValidFieldValue(enUnit, "bsd_maintenancefeespercent");
            #endregion

            newQuotation.Id = Guid.NewGuid();
            service.Create(newQuotation);
            return newQuotation.Id;
        }

        private EntityReference GetPhasesLaunch(Entity enUnit, ref Entity newQuotation)
        {
            traceService.Trace("GetPhasesLaunch");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_phaseslaunch"">
                <attribute name=""bsd_phaseslaunchid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""createdon"" />
                <attribute name=""bsd_depositamount"" />
                <attribute name=""bsd_minimumdeposit"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000000"" />
                  <condition attribute=""bsd_stopselling"" operator=""eq"" value=""0"" />
                </filter>
                <link-entity name=""bsd_bsd_phaseslaunch_bsd_product"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" intersect=""true"">
                  <filter>
                    <condition attribute=""bsd_productid"" operator=""eq"" value=""{enUnit.Id}"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count == 1)
            {
                Entity item = rs.Entities[0];
                newQuotation["bsd_depositfee"] = item.Contains("bsd_depositamount") ? item["bsd_depositamount"] : null;
                newQuotation["bsd_minimumdepositfee"] = item.Contains("bsd_minimumdeposit") ? item["bsd_minimumdeposit"] : null;
                return item.ToEntityReference();
            }
            return null;
        }
    }
}