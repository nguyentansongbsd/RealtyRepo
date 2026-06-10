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
                Entity enAppendix = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_type", "bsd_ra", "bsd_spa", "bsd_discountnew" }));
                if (!enAppendix.Contains("bsd_type"))
                    return;
                int bsd_type = ((OptionSetValue)enAppendix["bsd_type"]).Value;

                Entity enContract = null;
                EntityReference refUnit = null;
                if (bsd_type == 100000000 && enAppendix.Contains("bsd_ra"))   //Reservation Contract
                {
                    EntityReference refContract = (EntityReference)enAppendix["bsd_ra"];
                    enContract = service.Retrieve(refContract.LogicalName, refContract.Id, new ColumnSet(new string[] { "bsd_detailamount", "bsd_packagesellingamount", "bsd_promotion",
                    "bsd_landvaluededuction", "bsd_taxcode", "bsd_unitno" }));
                    refUnit = (EntityReference)enContract["bsd_unitno"];
                }
                else if (bsd_type == 100000001 && enAppendix.Contains("bsd_spa"))   //Option Entry
                {
                    EntityReference refContract = (EntityReference)enAppendix["bsd_spa"];
                    enContract = service.Retrieve(refContract.LogicalName, refContract.Id, new ColumnSet(new string[] { "bsd_detailamount", "bsd_packagesellingamount", "bsd_promotion",
                    "bsd_landvaluededuction", "bsd_taxcode", "bsd_unitnumber" }));
                    refUnit = (EntityReference)enContract["bsd_unitnumber"];
                }

                if (enContract == null)
                    return;

                decimal dicountNew = enAppendix.Contains("bsd_discountnew") ? ((Money)enAppendix["bsd_discountnew"]).Value : 0;

                decimal bsd_detailamount = enContract.Contains("bsd_detailamount") ? ((Money)enContract["bsd_detailamount"]).Value : 0;
                decimal bsd_packagesellingamount = enContract.Contains("bsd_packagesellingamount") ? ((Money)enContract["bsd_packagesellingamount"]).Value : 0;
                decimal bsd_promotion = enContract.Contains("bsd_promotion") ? ((Money)enContract["bsd_promotion"]).Value : 0;
                decimal bsd_landvaluededuction = enContract.Contains("bsd_landvaluededuction") ? ((Money)enContract["bsd_landvaluededuction"]).Value : 0;

                Entity upAppendix = new Entity(enAppendix.LogicalName, enAppendix.Id);
                //upAppendix["bsd_discountnew"] = new Money(dicountNew);
                upAppendix["bsd_promotionnew"] = new Money(bsd_promotion);
                upAppendix["bsd_packagesellingamountnew"] = new Money(bsd_packagesellingamount);
                upAppendix["bsd_landvaluedeductionnew"] = new Money(bsd_landvaluededuction);

                decimal bsd_totalamountlessfreightnew = bsd_detailamount - dicountNew - bsd_promotion + bsd_packagesellingamount;
                upAppendix["bsd_totalamountlessfreightnew"] = new Money(bsd_totalamountlessfreightnew);

                decimal bsd_totaltaxnew = GetTotalTax(enContract, bsd_totalamountlessfreightnew, bsd_landvaluededuction);
                upAppendix["bsd_totaltaxnew"] = new Money(bsd_totaltaxnew);

                decimal bsd_totalamountlessfreightvatnew = bsd_totalamountlessfreightnew + bsd_totaltaxnew;
                upAppendix["bsd_totalamountlessfreightvatnew"] = new Money(bsd_totalamountlessfreightvatnew);

                decimal bsd_maintenancefeesnew = GetMaintenanceFee(refUnit, bsd_totalamountlessfreightnew);
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

        private decimal GetTotalTax(Entity enContract, decimal bsd_totalamountlessfreight, decimal bsd_landvaluededuction)
        {
            traceService.Trace("GetTotalTax");

            decimal percentTax = 0;
            if (enContract.Contains("bsd_taxcode"))
            {
                EntityReference refTax = (EntityReference)enContract["bsd_taxcode"];
                Entity enTax = service.Retrieve(refTax.LogicalName, refTax.Id, new ColumnSet(new string[] { "bsd_value" }));
                percentTax = enTax.Contains("bsd_value") ? (decimal)enTax["bsd_value"] / 100 : 0;
            }
            decimal bsd_totaltax = (bsd_totalamountlessfreight - bsd_landvaluededuction) * percentTax;
            return bsd_totaltax;
        }

        private decimal GetMaintenanceFee(EntityReference refUnit, decimal bsd_totalamountlessfreight)
        {
            traceService.Trace("GetMaintenanceFee");

            Entity enUnit = service.Retrieve(refUnit.LogicalName, refUnit.Id, new ColumnSet(new string[] { "bsd_maintenancefeespercent" }));
            decimal bsd_maintenancefeespercent = enUnit.Contains("bsd_maintenancefeespercent") ? (decimal)enUnit["bsd_maintenancefeespercent"] / 100 : 0;
            return bsd_maintenancefeespercent * bsd_totalamountlessfreight;
        }
    }
}