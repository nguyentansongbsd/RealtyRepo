using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PSAppendix_CreateUpdate
{
    public class Plugin_PSAppendix_CreateUpdate : IPlugin
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
                Entity enAppendix = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_spa", "bsd_discountnew" }));
                EntityReference refSPA = (EntityReference)enAppendix["bsd_spa"];
                Entity enOE = service.Retrieve(refSPA.LogicalName, refSPA.Id, new ColumnSet(new string[] { "bsd_detailamount", "bsd_packagesellingamount", "bsd_promotion",
                    "bsd_landvaluededuction", "bsd_taxcode", "bsd_unitnumber" }));

                decimal dicountNew = enAppendix.Contains("bsd_discountnew") ? ((Money)enAppendix["bsd_discountnew"]).Value : 0;

                decimal bsd_detailamount = enOE.Contains("bsd_detailamount") ? ((Money)enOE["bsd_detailamount"]).Value : 0;
                decimal bsd_packagesellingamount = enOE.Contains("bsd_packagesellingamount") ? ((Money)enOE["bsd_packagesellingamount"]).Value : 0;
                decimal bsd_promotion = enOE.Contains("bsd_promotion") ? ((Money)enOE["bsd_promotion"]).Value : 0;
                decimal bsd_landvaluededuction = enOE.Contains("bsd_landvaluededuction") ? ((Money)enOE["bsd_landvaluededuction"]).Value : 0;

                Entity upAppendix = new Entity(enAppendix.LogicalName, enAppendix.Id);
                //upAppendix["bsd_discountnew"] = new Money(dicountNew);
                upAppendix["bsd_promotionnew"] = new Money(bsd_promotion);
                upAppendix["bsd_packagesellingamountnew"] = new Money(bsd_packagesellingamount);
                upAppendix["bsd_landvaluedeductionnew"] = new Money(bsd_landvaluededuction);

                decimal bsd_totalamountlessfreightnew = bsd_detailamount - dicountNew - bsd_promotion + bsd_packagesellingamount;
                upAppendix["bsd_totalamountlessfreightnew"] = new Money(bsd_totalamountlessfreightnew);

                decimal bsd_totaltaxnew = GetTotalTax(enOE, bsd_totalamountlessfreightnew, bsd_landvaluededuction);
                upAppendix["bsd_totaltaxnew"] = new Money(bsd_totaltaxnew);

                decimal bsd_totalamountlessfreightvatnew = bsd_totalamountlessfreightnew + bsd_totaltaxnew;
                upAppendix["bsd_totalamountlessfreightvatnew"] = new Money(bsd_totalamountlessfreightvatnew);

                decimal bsd_maintenancefeesnew = GetMaintenanceFee(enOE, bsd_totalamountlessfreightnew);
                upAppendix["bsd_maintenancefeesnew"] = new Money(bsd_maintenancefeesnew);

                upAppendix["bsd_totalamountnew"] = new Money(bsd_totalamountlessfreightvatnew + bsd_maintenancefeesnew);
                service.Update(upAppendix);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private decimal GetTotalTax(Entity enOE, decimal bsd_totalamountlessfreight, decimal bsd_landvaluededuction)
        {
            traceService.Trace("GetTotalTax");

            decimal percentTax = 0;
            if (enOE.Contains("bsd_taxcode"))
            {
                EntityReference refTax = (EntityReference)enOE["bsd_taxcode"];
                Entity enTax = service.Retrieve(refTax.LogicalName, refTax.Id, new ColumnSet(new string[] { "bsd_value" }));
                percentTax = enTax.Contains("bsd_value") ? (decimal)enTax["bsd_value"] / 100 : 0;
            }
            decimal bsd_totaltax = (bsd_totalamountlessfreight - bsd_landvaluededuction) * percentTax;
            return bsd_totaltax;
        }

        private decimal GetMaintenanceFee(Entity enOE, decimal bsd_totalamountlessfreight)
        {
            traceService.Trace("GetMaintenanceFee");

            EntityReference refUnit = (EntityReference)enOE["bsd_unitnumber"];
            Entity enUnit = service.Retrieve(refUnit.LogicalName, refUnit.Id, new ColumnSet(new string[] { "bsd_maintenancefeespercent" }));
            decimal bsd_maintenancefeespercent = enUnit.Contains("bsd_maintenancefeespercent") ? (decimal)enUnit["bsd_maintenancefeespercent"] / 100 : 0;
            return bsd_maintenancefeespercent * bsd_totalamountlessfreight;
        }
    }
}