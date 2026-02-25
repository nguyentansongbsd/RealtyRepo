using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

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
                "bsd_landvalueofunit"}));
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
            decimal bsd_netsaleablearea = enUnit.Contains("bsd_netsaleablearea") ? (decimal)enUnit["bsd_netsaleablearea"] : 0;
            newOE["bsd_netusablearea"] = bsd_netsaleablearea;

            decimal bsd_landvalueofunit = enUnit.Contains("bsd_landvalueofunit") ? ((Money)enUnit["bsd_landvalueofunit"]).Value : 0;
            newOE["bsd_landvaluededuction"] = new Money(bsd_landvalueofunit * bsd_netsaleablearea);

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
            newOE["bsd_detailamount"] = enUnit.Contains("bsd_price") ? enUnit["bsd_price"] : new Money(0);
            newOE["bsd_totalamountlessfreight"] = enUnit.Contains("bsd_price") ? enUnit["bsd_price"] : new Money(0);
            newOE["bsd_discount"] = new Money(0);

            newOE.Id = Guid.NewGuid();
            service.Create(newOE);
            create_update_DataProjection(enUnit.Id, newOE);
            return newOE.Id;
        }
        private void create_update_DataProjection(Guid idUnit, Entity enEntity)
        {
            // get DataProjection theo unit
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_dataprojection"">
                <filter>
                  <condition attribute=""bsd_productid"" operator=""eq"" value=""{idUnit}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection en = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (en.Entities.Count > 0)
            {
                Entity enDataprojection = en.Entities[0];
                Entity enUp = new Entity(enDataprojection.LogicalName, enDataprojection.Id);
                enUp["bsd_spaid"] = enEntity.ToEntityReference();
                if (enEntity.Contains("bsd_customerid")) enUp["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_project")) enUp["bsd_project"] = enEntity["bsd_project"];
                if (enEntity.Contains("bsd_opportunityid")) enUp["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_phaseslaunch")) enUp["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunch"];
                if (enEntity.Contains("bsd_reservationcontract")) enUp["bsd_raid"] = enEntity["bsd_reservationcontract"];
                if (enEntity.Contains("bsd_quoteid")) enUp["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Update(enUp);
            }
            else
            {
                Entity enCre = new Entity("bsd_dataprojection");
                enCre["bsd_spaid"] = enEntity.ToEntityReference();
                if (enEntity.Contains("bsd_customerid")) enCre["bsd_customerid"] = enEntity["bsd_customerid"];
                if (enEntity.Contains("bsd_project")) enCre["bsd_project"] = enEntity["bsd_project"];
                if (enEntity.Contains("bsd_opportunityid")) enCre["bsd_bookingid"] = enEntity["bsd_opportunityid"];
                if (enEntity.Contains("bsd_phaseslaunch")) enCre["bsd_phaselaunchid"] = enEntity["bsd_phaseslaunch"];
                if (enEntity.Contains("bsd_unitno")) enCre["bsd_productid"] = enEntity["bsd_unitno"];
                if (enEntity.Contains("bsd_reservationcontract")) enCre["bsd_raid"] = enEntity["bsd_reservationcontract"];
                if (enEntity.Contains("bsd_quoteid")) enCre["bsd_depositid"] = enEntity["bsd_quoteid"];
                service.Create(enCre);
            }
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