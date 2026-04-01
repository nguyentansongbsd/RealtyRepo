using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Security.Policy;

namespace Action_AreaAppendixy_Approved
{
    public class Action_AreaAppendixy_Approved : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService traceS = null;
        EntityReference target = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                target = (EntityReference)context.InputParameters["Target"];
                traceS = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceS.Trace($"start {target.Id}");
                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                traceS.Trace("2");
                if (enTarget.Contains("bsd_optionentry"))
                {
                    traceS.Trace("2");
                    EntityReference refOE = (EntityReference)enTarget["bsd_optionentry"];
                    deletePaymentSchemeDetail(refOE);
                    mapPaymentSchemeDetail(refOE);
                    Entity enSPA = new Entity(refOE.LogicalName, refOE.Id);
                    enSPA["bsd_totalamountlessfreight"] = new Money(enTarget.Contains("bsd_bsd_totalamountlessfreightnew") ? ((Money)enTarget["bsd_bsd_totalamountlessfreightnew"]).Value : 0);
                    enSPA["bsd_totaltax"] = new Money(enTarget.Contains("bsd_totaltaxnew") ? ((Money)enTarget["bsd_totaltaxnew"]).Value : 0);
                    enSPA["bsd_freightamount"] = new Money(enTarget.Contains("bsd_maintenancefeesnew") ? ((Money)enTarget["bsd_maintenancefeesnew"]).Value : 0);
                    enSPA["bsd_totalamountlessfreightaftervat"] = new Money(enTarget.Contains("bsd_totalamountlessfreightvatnew") ? ((Money)enTarget["bsd_totalamountlessfreightvatnew"]).Value : 0);
                    enSPA["bsd_totalamount"] = new Money(enTarget.Contains("bsd_totalamountnew") ? ((Money)enTarget["bsd_totalamountnew"]).Value : 0);
                    service.Update(enSPA);
                }

                Entity enUp = new Entity(target.LogicalName, target.Id);
                enUp["statuscode"] = new OptionSetValue(100000002);
                enUp["bsd_confirmedby"] = new EntityReference("systemuser", context.UserId);
                enUp["bsd_confirmeddate"] = DateTime.Now;
                service.Update(enUp);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void deletePaymentSchemeDetail(EntityReference refOE)
        {
            traceS.Trace("delete_PaymentSchemeDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""bsd_optionentry"" operator=""eq"" value=""{refOE.Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (Entity item in rs.Entities)
                {
                    service.Delete(item.LogicalName, item.Id);
                }
            }
        }
        private void mapPaymentSchemeDetail(EntityReference refOE)
        {
            traceS.Trace("genPaymentSchemeDetail");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_areaappendix"" operator=""eq"" value=""{target.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    CreateNewFromItem(item, refOE);
                }
            }
        }
        private void CreateNewFromItem(Entity item, EntityReference refOE)
        {
            Entity it = new Entity(item.LogicalName);
            it = item;
            it.Attributes.Remove(item.LogicalName + "id");
            it.Attributes.Remove("ownerid");
            it.Attributes.Remove("bsd_areaappendix");
            it["bsd_optionentry"] = refOE;
            it.Id = Guid.NewGuid();
            service.Create(it);
        }
    }
}