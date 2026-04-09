using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Floor_FloorCode
{
    public class Plugin_Floor_FloorCode : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationService service = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private ITracingService tracingService = null;

        Entity enFloor = null;
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
                this.enFloor = this.service.Retrieve(target.LogicalName, target.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_floor", "bsd_block"));

                UpdateFloorCode();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        private void UpdateFloorCode()
        {
            try
            {
                tracingService.Trace("Start update floor code");
                string blockCode = getBlockCode(enFloor.GetAttributeValue<EntityReference>("bsd_block"));
                string floorName = enFloor.GetAttributeValue<string>("bsd_floor");
                string FloorCode = blockCode + floorName;
                Entity enFloor_up = new Entity(enFloor.LogicalName, enFloor.Id);
                enFloor_up["bsd_name"] = FloorCode;
                service.Update(enFloor_up);
                tracingService.Trace("End update floor code");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private string getBlockCode(EntityReference enfBlock)
        {
            try
            {
                tracingService.Trace("Get block code");
                Entity enProject = service.Retrieve(enfBlock.LogicalName, enfBlock.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_name"));
                if (enProject.Contains("bsd_name"))
                    return enProject["bsd_name"].ToString() + "-";
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
