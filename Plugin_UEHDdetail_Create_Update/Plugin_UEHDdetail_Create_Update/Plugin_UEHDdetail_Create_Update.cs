using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_UEHDdetail_Create_Update
{
    public class Plugin_UEHDdetail_Create_Update : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            ITracingService traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.MessageName == "Create")
            {
                if (context.Depth != 2) return;
                Entity target = (Entity)context.InputParameters["Target"];
                Entity enUp = new Entity(target.LogicalName, target.Id);
                if (target.Contains("bsd_updateestimatehandoverdate"))
                {
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_updateestimatehandoverdate"">
                            <attribute name=""bsd_project"" />
                            <filter>
                              <condition attribute=""bsd_updateestimatehandoverdateid"" operator=""eq"" value=""{((EntityReference)target["bsd_updateestimatehandoverdate"]).Id}"" />
                              <condition attribute=""bsd_project"" operator=""not-null"" />
                            </filter>
                          </entity>
                        </fetch>";
                    EntityCollection enMater = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (enMater.Entities.Count == 0) throw new InvalidPluginExecutionException("Update Estimate Handover Date not found.");
                    foreach (Entity entity in enMater.Entities)
                    {
                        EntityReference bsd_project = (EntityReference)entity["bsd_project"];
                        enUp["bsd_project"] = bsd_project;
                        if (target.Contains("bsd_units"))
                        {
                            fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch>
                              <entity name=""bsd_salesorder"">
                                <attribute name=""bsd_salesorderid"" />
                                <filter>
                                  <condition attribute=""bsd_unitnumber"" operator=""eq"" value=""{((EntityReference)target["bsd_units"]).Id}"" />
                                  <condition attribute=""bsd_project"" operator=""eq"" value=""{bsd_project.Id}"" />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection enSPA = service.RetrieveMultiple(new FetchExpression(fetchXml));
                            if (enSPA.Entities.Count == 0) throw new InvalidPluginExecutionException("SPA not found.");
                            foreach (Entity entity2 in enSPA.Entities)
                            {
                                enUp["bsd_optionentry"] = entity2.ToEntityReference();
                                fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                                <fetch>
                                  <entity name=""bsd_paymentschemedetail"">
                                    <attribute name=""bsd_name"" />
                                    <attribute name=""bsd_paymentschemedetailid"" />
                                    <attribute name=""bsd_duedate"" />
                                    <filter>
                                      <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{entity2.Id}"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""{0}"" />
                                      <condition attribute=""bsd_pinkbookhandover"" operator=""eq"" value=""{1}"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                                EntityCollection enInstallment = service.RetrieveMultiple(new FetchExpression(fetchXml));
                                if (enInstallment.Entities.Count == 0) throw new InvalidPluginExecutionException("Installment not found.");
                                foreach (Entity entity3 in enInstallment.Entities)
                                {
                                    enUp["bsd_installment"] = entity3.ToEntityReference();
                                    enUp["bsd_estimatehandoverdateold"] = entity3.Contains("bsd_duedate") ? entity3["bsd_duedate"] : null;
                                    service.Update(enUp);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
