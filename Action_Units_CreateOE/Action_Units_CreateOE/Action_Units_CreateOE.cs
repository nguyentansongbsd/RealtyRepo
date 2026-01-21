using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Units_CreateOE
{
    public class Action_Units_CreateOE : IPlugin
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
                "bsd_maintenancefeespercent"}));
                int status = enUnit.Contains("statuscode") ? ((OptionSetValue)enUnit["statuscode"]).Value : -99;
                if (status != 100000000) //Available
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "invalid_status_unit"));

                Guid idOE = CreateOE(enUnit, target);
                UpdateUnit(enUnit);

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
            return enUnit.Contains(field) ? enUnit[field] : null; ;
        }

        private Guid CreateOE(Entity enUnit, EntityReference target)
        {
            traceService.Trace("CreateOE");

            Entity newOE = new Entity("bsd_salesorder");
            newOE["bsd_name"] = GetValidFieldValue(enUnit, "bsd_name");
            newOE["bsd_date"] = DateTime.UtcNow;
            newOE["bsd_project"] = GetValidFieldValue(enUnit, "bsd_projectcode");
            newOE["bsd_phaseslaunch"] = GetPhasesLaunch(enUnit);
            //newOE["bsd_pricelevel"] = GetValidFieldValue(enUnit, "bsd_pricelevel");
            newOE["bsd_taxcode"] = GetValidFieldValue(enUnit, "bsd_taxcode");
            newOE["bsd_unittype"] = GetValidFieldValue(enUnit, "bsd_unittype");
            newOE["bsd_unitnumber"] = target;
            decimal bsd_netsaleablearea = 0;
            if (enUnit.Contains("bsd_netsaleablearea"))
            {
                bsd_netsaleablearea = (decimal)enUnit["bsd_netsaleablearea"];
                newOE["bsd_netusablearea"] = bsd_netsaleablearea;
            }

            //Management Fee Information
            #region Management Fee Information
            int bsd_numberofmonthspaidmf = 0;
            if (enUnit.Contains("bsd_numberofmonthspaidmf"))
            {
                bsd_numberofmonthspaidmf = (int)enUnit["bsd_numberofmonthspaidmf"];
                newOE["bsd_numberofmonthspaidmf"] = bsd_numberofmonthspaidmf;
            }
            decimal bsd_managementamountmonth = enUnit.Contains("bsd_managementamountmonth") ? ((Money)enUnit["bsd_managementamountmonth"]).Value : 0;
            newOE["bsd_managementfee"] = new Money(bsd_netsaleablearea * bsd_managementamountmonth * bsd_numberofmonthspaidmf);
            #endregion

            //Price
            #region Price
            decimal bsd_totalamountlessfreight = 0;
            if (enUnit.Contains("bsd_price"))
            {
                bsd_totalamountlessfreight = ((Money)enUnit["bsd_price"]).Value;
                newOE["bsd_detailamount"] = new Money(bsd_totalamountlessfreight);
                newOE["bsd_totalamountlessfreight"] = new Money(bsd_totalamountlessfreight);
            }
            decimal bsd_landvaluededuction = 0;
            decimal percentTax = 0;
            if (enUnit.Contains("bsd_taxcode"))
            {
                EntityReference refTax = (EntityReference)enUnit["bsd_taxcode"];
                Entity enTax = service.Retrieve(refTax.LogicalName, refTax.Id, new ColumnSet(new string[] { "bsd_value" }));
                percentTax = enTax.Contains("bsd_value") ? (decimal)enTax["bsd_value"] / 100 : 0;
            }
            decimal bsd_totaltax = (bsd_totalamountlessfreight - bsd_landvaluededuction) * percentTax;
            newOE["bsd_totaltax"] = new Money(bsd_totaltax);
            decimal bsd_maintenancefeespercent = enUnit.Contains("bsd_maintenancefeespercent") ? (decimal)enUnit["bsd_maintenancefeespercent"] : 0;
            decimal bsd_freightamount = bsd_maintenancefeespercent * bsd_totalamountlessfreight;
            newOE["bsd_freightamount"] = new Money(bsd_freightamount);
            newOE["bsd_totalamount"] = new Money(bsd_totalamountlessfreight + bsd_totaltax + bsd_freightamount);
            #endregion

            newOE.Id = Guid.NewGuid();
            service.Create(newOE);

            return newOE.Id;
        }

        private void UpdateUnit(Entity enUnit)
        {
            traceService.Trace("UpdateUnit");

            Entity upUnit = new Entity(enUnit.LogicalName, enUnit.Id);
            upUnit["statuscode"] = new OptionSetValue(100000008);    //In Contract
            service.Update(upUnit);
        }

        private EntityReference GetPhasesLaunch(Entity enUnit)
        {
            traceService.Trace("GetPhasesLaunch");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_phaseslaunch"">
                <attribute name=""bsd_phaseslaunchid"" />
                <attribute name=""bsd_name"" />
                <attribute name=""createdon"" />
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
                return rs.Entities[0].ToEntityReference();
            }
            return null;
        }
    }
}