using Microsoft.Xrm.Sdk;
using System;

namespace Plugin_Discount_Create_Update
{
    public class Plugin_Discount_Create_Update : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.Depth > 2) return;
            if (!context.InputParameters.Contains("Target")) return;

            var target = (Entity)context.InputParameters["Target"];
            if (!target.Contains("bsd_description"))
            {
                target["bsd_fetchxml"] = null;
            }
            else
            {
                var description = target.GetAttributeValue<string>("bsd_description");
                if (string.IsNullOrWhiteSpace(description))
                {
                    target["bsd_fetchxml"] = null;
                }
                else
                {
                    var rootEntity = QueryParser.ExtractRootEntity(description);
                    var fetchXml = QueryParser.Convert(description, rootEntity);
                    target["bsd_fetchxml"] = fetchXml;
                }
            }
        }
    }
}
