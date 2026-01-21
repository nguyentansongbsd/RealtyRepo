using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Management;

namespace Plugin_Product_UnitCodeSystem
{
    public class Plugin_Product_UnitCodeSystem : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationService service = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private ITracingService tracingService = null;

        Entity enUnit = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this.context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this.tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            this.serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this.service = serviceFactory.CreateOrganizationService(this.context.UserId);

            try
            {
                if (this.context.Depth > 3) return;
                if (this.context.MessageName != "Create" && this.context.MessageName != "Update") return;
                var target = (Entity)this.context.InputParameters["Target"];
                this.enUnit = this.service.Retrieve(target.LogicalName, target.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_name", "bsd_floor"));

                UpdateUnitCodeSystem();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void UpdateUnitCodeSystem()
        {
            try
            {
                tracingService.Trace("Start update unit code system");
                string floorCode = getFloorCode(enUnit.GetAttributeValue<EntityReference>("bsd_floor"));
                string unitCode = enUnit.GetAttributeValue<string>("bsd_name");
                string unitCodeSystem = floorCode + unitCode;
                Entity enUnit_Up = new Entity(enUnit.LogicalName, enUnit.Id);
                enUnit_Up["bsd_unitcodesystem"] = unitCodeSystem;
                service.Update(enUnit_Up);
                tracingService.Trace("End update unit code system");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private string getFloorCode(EntityReference enfFloor)
        {
            try
            {
                tracingService.Trace("Get Floor code");
                Entity enFloor = service.Retrieve(enfFloor.LogicalName, enfFloor.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_name"));
                if(enFloor.Contains("bsd_name"))
                    return enFloor["bsd_name"].ToString() + "-";
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
