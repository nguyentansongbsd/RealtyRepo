using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PhasesLaunch_Approve
{
    public class Plugin_PhasesLaunch_Approve : IPlugin
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
                Entity enPL = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode" }));
                int status = enPL.Contains("statuscode") ? ((OptionSetValue)enPL["statuscode"]).Value : -99;
                if (status != 100000000)  //Launched
                    return;

                var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch distinct=""true"">
                  <entity name=""bsd_phaseslaunch"">
                    <attribute name=""bsd_phaseslaunchid"" />
                    <attribute name=""bsd_name"" />
                    <filter>
                      <condition attribute=""statuscode"" operator=""eq"" value=""100000000"" />
                      <condition attribute=""bsd_phaseslaunchid"" operator=""ne"" value=""{enPL.Id}"" />
                    </filter>
                    <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_phaseslaunchid"" to=""bsd_phaseslaunchid"" alias=""bsd_bsd_phaseslaunch_bsd_pricelevel"" intersect=""true"">
                      <link-entity name=""bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                        <link-entity name=""bsd_bsd_phaseslaunch_bsd_pricelevel"" from=""bsd_pricelevelid"" to=""bsd_pricelevelid"" intersect=""true"">
                          <filter>
                            <condition attribute=""bsd_phaseslaunchid"" operator=""eq"" value=""{enPL.Id}"" />
                          </filter>
                        </link-entity>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                {
                    string name = rs.Entities[0].Contains("bsd_name") ? (string)rs.Entities[0]["bsd_name"] : string.Empty;
                    traceService.Trace($"bsd_phaseslaunch: {name} {rs.Entities[0].Id}");
                    throw new InvalidPluginExecutionException($"This price list has been launched under the phase launch '{name}'. Please check the information.");
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}