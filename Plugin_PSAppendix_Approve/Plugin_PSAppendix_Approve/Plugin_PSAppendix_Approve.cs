using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_PSAppendix_Approve
{
    public class Plugin_PSAppendix_Approve : IPlugin
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
                Entity enAppendix = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                int status = enAppendix.Contains("statuscode") ? ((OptionSetValue)enAppendix["statuscode"]).Value : -99;
                if (status != 100000001)  //Approved
                    return;

                if (!enAppendix.Contains("bsd_type"))
                    return;
                int bsd_type = ((OptionSetValue)enAppendix["bsd_type"]).Value;

                EntityReference refContract = null;
                string logicalName = string.Empty;
                if (bsd_type == 100000000 && enAppendix.Contains("bsd_ra"))   //Reservation Contract
                {
                    refContract = (EntityReference)enAppendix["bsd_ra"];
                    logicalName = "bsd_reservationcontract";
                }
                else if (bsd_type == 100000001 && enAppendix.Contains("bsd_spa"))   //Option Entry
                {
                    refContract = (EntityReference)enAppendix["bsd_spa"];
                    logicalName = "bsd_optionentry";
                }

                if (refContract == null)
                    return;

                Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
                if (refContract.LogicalName == "bsd_reservationcontract")
                {
                    upContract["bsd_discountamount"] = ValidValue(enAppendix, "bsd_discountnew");
                }
                else
                {
                    upContract["bsd_discount"] = ValidValue(enAppendix, "bsd_discountnew");
                }

                upContract["bsd_promotion"] = ValidValue(enAppendix, "bsd_promotionnew");
                upContract["bsd_packagesellingamount"] = ValidValue(enAppendix, "bsd_packagesellingamountnew");
                upContract["bsd_totalamountlessfreight"] = ValidValue(enAppendix, "bsd_totalamountlessfreightnew");
                upContract["bsd_landvaluededuction"] = ValidValue(enAppendix, "bsd_landvaluedeductionnew");
                upContract["bsd_totaltax"] = ValidValue(enAppendix, "bsd_totaltaxnew");
                upContract["bsd_freightamount"] = ValidValue(enAppendix, "bsd_maintenancefeesnew");
                upContract["bsd_totalamountlessfreightaftervat"] = ValidValue(enAppendix, "bsd_totalamountlessfreightvatnew");
                upContract["bsd_totalamount"] = ValidValue(enAppendix, "bsd_totalamountnew");
                upContract["bsd_paymentscheme"] = ValidValue(enAppendix, "bsd_paymentschemenew");

                service.Update(upContract);

                DeleteInstallment(refContract, logicalName);
                CreateInstallment(enAppendix, refContract, logicalName);

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private object ValidValue(Entity enAppendix, string field)
        {
            return enAppendix.Contains(field) ? enAppendix[field] : null;
        }


        private void DeleteInstallment(EntityReference refContract, string logicalName)
        {
            traceService.Trace("vào DeleteInstallment");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <attribute name=""bsd_paymentschemedetailid"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""{logicalName}"" operator=""eq"" value=""{refContract.Id}"" />
                </filter>
              </entity>
            </fetch>";
            var rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (Entity item in rs.Entities)
                {
                    service.Delete(item.LogicalName, item.Id);
                }
            }
        }

        private void CreateInstallment(Entity enAppendix, EntityReference refContract, string logicalName)
        {
            traceService.Trace("vào CreateInstallment");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_psappendix"" operator=""eq"" value=""{enAppendix.Id}"" />
                </filter>
              </entity>
            </fetch>";
            var rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                foreach (Entity item in rs.Entities)
                {
                    Entity it = new Entity(item.LogicalName);
                    item.Attributes.Remove(item.LogicalName + "id");
                    item.Attributes.Remove("ownerid");
                    item.Attributes.Remove(enAppendix.LogicalName);
                    item[logicalName] = refContract;
                    item.Id = Guid.NewGuid();
                    it = item;
                    service.Create(it);
                }
            }
        }
    }
}