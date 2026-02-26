using Microsoft.Xrm.Sdk;
using System;
using Plugin_QueryBuilderGroup_Create_Update.Services;
using Microsoft.Xrm.Sdk.Query;

namespace Plugin_QueryBuilderGroup_Create_Update
{
    public class Plugin_QueryBuilderGroup_Create_Update : IPlugin
    {
        public void Execute(IServiceProvider sp)
        {
            var context =
                (IPluginExecutionContext)
                sp.GetService(typeof(IPluginExecutionContext));
            ITracingService traceService = (ITracingService)sp.GetService(typeof(ITracingService));
            traceService.Trace("Plugin_QueryBuilderGroup_Create_Update");
            traceService.Trace("Depth\n" + context.Depth);
            if (context.Depth > 2) return;
            if (!context.InputParameters.Contains("Target"))
                return;

            var entity =
                (Entity)context.InputParameters["Target"];

            if (!entity.Contains("bsd_jsonconfig"))
                return;
            
            string json =
                entity.GetAttributeValue<string>(
                    "bsd_jsonconfig");

            string parentEntity =
                entity.GetAttributeValue<string>(
                    "bsd_parententity");
            if (parentEntity == null)
            {
                IOrganizationServiceFactory factory = factory = (IOrganizationServiceFactory)sp.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = service = factory.CreateOrganizationService(context.UserId);
                Entity enTarget = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                parentEntity = enTarget.Contains("bsd_parententity") ? (string)enTarget["bsd_parententity"] : "";
            }
            traceService.Trace("json\n" + json);
            traceService.Trace("parentEntity\n" + parentEntity);
            var parser = new JsonQueryParser();
            var tree = parser.Parse(json);

            var compiler = new FetchXmlCompiler();

            string fetch =
                compiler.Build(tree, parentEntity);

            entity["bsd_fetchxml"] = fetch;
        }
    }
}