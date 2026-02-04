using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_SubSale_Approved
{
    public class Plugin_SubSale_Approved : IPlugin
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
                Entity enSubSale = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_type", "bsd_reservation",
                    "bsd_reservationcontract", "bsd_optionentry", "bsd_newcustomer" }));
                int status = enSubSale.Contains("statuscode") ? ((OptionSetValue)enSubSale["statuscode"]).Value : -99;
                if (status != 100000002 || !enSubSale.Contains("bsd_type"))  //Complete
                    return;

                int bsd_type = ((OptionSetValue)enSubSale["bsd_type"]).Value;
                string logicalName = null;
                if (bsd_type == 100000000 && enSubSale.Contains("bsd_reservation"))   //Deposit
                {
                    logicalName = "bsd_reservation";
                }
                else if (bsd_type == 100000001 && enSubSale.Contains("bsd_reservationcontract"))   //Reservation Contract
                {
                    logicalName = "bsd_reservationcontract";
                }
                else if (bsd_type == 100000002 && enSubSale.Contains("bsd_optionentry"))   //Option Entry
                {
                    logicalName = "bsd_optionentry";
                }

                if (logicalName == null)
                    return;
                EntityReference refContract = (EntityReference)enSubSale[logicalName];

                UpdateContract(refContract, logicalName, enSubSale);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void UpdateContract(EntityReference refContract, string logicalName, Entity enSubSale)
        {
            traceService.Trace("UpdateContract");

            Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
            upContract["bsd_customerid"] = enSubSale.Contains("bsd_newcustomer") ? enSubSale["bsd_newcustomer"] : null;
            service.Update(upContract);

            DeleteOldCoOwner(refContract, logicalName);
            CreateNewCoOwner(refContract, logicalName, enSubSale);
        }

        private void DeleteOldCoOwner(EntityReference refContract, string logicalName)
        {
            traceService.Trace("DeleteOldCoOwner");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_coowner"">
                    <filter>
                      <condition attribute=""{logicalName}"" operator=""eq"" value=""{refContract.Id}"" />
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                    </filter>
                    <order attribute=""createdon"" />
                  </entity>
                </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    service.Delete(item.LogicalName, item.Id);
                }
            }
        }

        private void CreateNewCoOwner(EntityReference refContract, string logicalName, Entity enSubSale)
        {
            traceService.Trace("CreateNewCoOwner");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                <fetch>
                  <entity name=""bsd_coowner"">
                    <filter>
                      <condition attribute=""bsd_assign"" operator=""eq"" value=""{enSubSale.Id}"" />
                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                    </filter>
                    <order attribute=""createdon"" />
                  </entity>
                </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    Entity newCoOwner = new Entity("bsd_coowner");
                    newCoOwner = item;
                    newCoOwner.Attributes.Remove("bsd_coownerid");
                    newCoOwner.Attributes.Remove("ownerid");
                    newCoOwner.Attributes.Remove("bsd_assign");
                    newCoOwner[logicalName] = refContract;
                    newCoOwner.Id = Guid.NewGuid();
                    service.Create(newCoOwner);
                }
            }
        }
    }
}