using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Block_BlockCode
{
    public class Plugin_Block_BlockCode : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationService service = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private ITracingService tracingService = null;

        Entity enBlock = null;
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
                this.enBlock = this.service.Retrieve(target.LogicalName, target.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_name", "bsd_project"));

                UpdateBlockCode();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        private void UpdateBlockCode()
        {
            try
            {
                tracingService.Trace("Start update block code");
                string projectCode = getProjectCode(enBlock.GetAttributeValue<EntityReference>("bsd_project"));
                string blockName = enBlock.GetAttributeValue<string>("bsd_name");
                string blockCode = projectCode + blockName;
                Entity enBlock_up = new Entity(enBlock.LogicalName, enBlock.Id);
                enBlock_up["bsd_blockcode"] = blockCode;
                service.Update(enBlock_up);
                tracingService.Trace("End update block code");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private string getProjectCode(EntityReference enfProject)
        {
            try
            {
                tracingService.Trace("Get project code");
                Entity enProject = service.Retrieve(enfProject.LogicalName, enfProject.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_projectcode"));
                if (enProject.Contains("bsd_projectcode"))
                    return enProject["bsd_projectcode"].ToString() + "-";
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

    }
}
