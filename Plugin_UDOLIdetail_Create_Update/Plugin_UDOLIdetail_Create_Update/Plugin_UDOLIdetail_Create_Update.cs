using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_UDOLIdetail_Create_Update
{
    public class Plugin_UDOLIdetail_Create_Update : IPlugin
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
                if (target.Contains("bsd_updateduedateoflastinstallment"))
                {
                    var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_masterupdateduedateoflastinstallment"">
                            <attribute name=""bsd_project"" />
                            <filter>
                              <condition attribute=""bsd_masterupdateduedateoflastinstallmentid"" operator=""eq"" value=""{((EntityReference)target["bsd_updateduedateoflastinstallment"]).Id}"" />
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
                        if (target.Contains("bsd_spa"))
                        {
                            fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch>
                              <entity name=""bsd_salesorder"">
                                <attribute name=""bsd_unitnumber"" />
                                <filter>
                                  <condition attribute=""bsd_salesorderid"" operator=""eq"" value=""{((EntityReference)target["bsd_spa"]).Id}"" />
                                  <condition attribute=""bsd_project"" operator=""eq"" value=""{bsd_project.Id}"" />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection enSPA = service.RetrieveMultiple(new FetchExpression(fetchXml));
                            if (enSPA.Entities.Count == 0) throw new InvalidPluginExecutionException("SPA not found.");
                            foreach (Entity entity2 in enSPA.Entities)
                            {
                                enUp["bsd_unit"] = (EntityReference)entity2["bsd_unitnumber"];
                                fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                                <fetch>
                                  <entity name=""bsd_paymentschemedetail"">
                                    <attribute name=""bsd_name"" />
                                    <attribute name=""bsd_paymentschemedetailid"" />
                                    <attribute name=""bsd_duedate"" />
                                    <filter>
                                      <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{entity2.Id}"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""{0}"" />
                                      <condition attribute=""bsd_lastinstallment"" operator=""eq"" value=""{1}"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                                EntityCollection enInstallment = service.RetrieveMultiple(new FetchExpression(fetchXml));
                                if (enInstallment.Entities.Count == 0) throw new InvalidPluginExecutionException("Installment not found.");
                                foreach (Entity entity3 in enInstallment.Entities)
                                {
                                    enUp["bsd_installment"] = entity3.ToEntityReference();
                                    enUp["bsd_duedateold"] = entity3.Contains("bsd_duedate") ? entity3["bsd_duedate"] : null;
                                    service.Update(enUp);
                                }
                            }
                        }
                    }
                }
            }
            else if (context.MessageName == "Update")
            {
                if (context.Depth > 2) return;
                Entity target = (Entity)context.InputParameters["Target"];
                int statuscode = ((OptionSetValue)target["statuscode"]).Value;
                if (statuscode == 667980001)//aprove
                {
                    Entity enDetail = service.Retrieve(target.LogicalName, target.Id, new ColumnSet("bsd_spa", "bsd_installment", "bsd_duedatenew"));
                    if (enDetail.Contains("bsd_spa"))
                    {
                        var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                            <fetch>
                              <entity name=""bsd_salesorder"">
                                <attribute name=""bsd_unitnumber"" />
                                <filter>
                                  <condition attribute=""bsd_salesorderid"" operator=""eq"" value=""{((EntityReference)enDetail["bsd_spa"]).Id}"" />
                                  <condition attribute=""statuscode"" operator=""eq"" value=""100000015"" />
                                </filter>
                              </entity>
                            </fetch>";
                        EntityCollection enSPA = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (enSPA.Entities.Count == 0) throw new InvalidPluginExecutionException("The SPA is invalid. Please check again.");
                    }
                    if (enDetail.Contains("bsd_installment") && enDetail.Contains("bsd_duedatenew"))
                    {
                        Entity enUp = new Entity(((EntityReference)enDetail["bsd_installment"]).LogicalName, ((EntityReference)enDetail["bsd_installment"]).Id);
                        enUp["bsd_duedate"] = RetrieveLocalTimeFromUTCTime((DateTime)enDetail["bsd_duedatenew"], service);
                        service.Update(enUp);
                    }
                }
                else return;
            }
        }
        private DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService service)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };
            var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);

            return response.LocalTime;
        }
        private int? RetrieveCurrentUsersSettings(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("systemuserid", ConditionOperator.EqualUserId) }
                }
            }).Entities[0].ToEntity<Entity>();
            return (int?)currentUserSettings.Attributes["timezonecode"];
        }
    }
}
