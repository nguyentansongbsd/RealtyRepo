using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Termination_Complete
{
    public class Action_Termination_Complete : IPlugin
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

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity enTermination = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "bsd_source", "bsd_reservation",
                    "bsd_reservationcontract", "bsd_optionentry", "bsd_units", "bsd_source", "bsd_resell", "bsd_customer", "bsd_project", "bsd_totalamountpaid", "bsd_forfeitureamount"}));

                int bsd_source = ((OptionSetValue)enTermination["bsd_source"]).Value;
                if (bsd_source == 100000000 && enTermination.Contains("bsd_reservation"))   //Deposit
                {
                    RunUpdate("bsd_reservation", enTermination, 667980004);
                }
                else if (bsd_source == 100000001 && enTermination.Contains("bsd_reservationcontract"))   //Reservation Contract
                {
                    RunUpdate("bsd_reservationcontract", enTermination, 100000004);
                }
                else if (bsd_source == 100000002 && enTermination.Contains("bsd_optionentry"))   //Option Entry
                {
                    RunUpdate("bsd_optionentry", enTermination, 100000014);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void RunUpdate(string logicalName, Entity enTermination, int statusContract)
        {
            traceService.Trace("RunUpdate");

            EntityReference refContract = (EntityReference)enTermination[logicalName];
            UpdateTermination(enTermination);
            UpdateContract(refContract, statusContract);

            EntityReference refUnit = (EntityReference)enTermination["bsd_units"];
            UpdateUnit(enTermination, refUnit);

            decimal bsd_forfeitureamount = enTermination.Contains("bsd_forfeitureamount") ? ((Money)enTermination["bsd_forfeitureamount"]).Value : 0;
            if (bsd_forfeitureamount > 0)
                CreateRefund(enTermination, refUnit, refContract, logicalName);
        }

        private void UpdateTermination(Entity enTermination)
        {
            traceService.Trace("UpdateTermination");

            Entity upTermination = new Entity(enTermination.LogicalName, enTermination.Id);
            upTermination["statecode"] = new OptionSetValue(1);    //inactive
            upTermination["statuscode"] = new OptionSetValue(100000005);    //Complete
            service.Update(upTermination);
        }

        private void UpdateContract(EntityReference refContract, int statusContract)
        {
            traceService.Trace("UpdateContract");

            Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
            upContract["statecode"] = new OptionSetValue(1);    //active
            upContract["statuscode"] = new OptionSetValue(statusContract);  //Terminated
            service.Update(upContract);
        }

        private void UpdateUnit(Entity enTermination, EntityReference refUnit)
        {
            traceService.Trace("UpdateUnit");

            bool bsd_resell = enTermination.Contains("bsd_resell") ? (bool)enTermination["bsd_resell"] : false;

            Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
            upUnit["statuscode"] = new OptionSetValue(bsd_resell ? 100000000 : 1);  //Available, Preparing
            service.Update(upUnit);
        }

        private void CreateRefund(Entity enTermination, EntityReference refUnit, EntityReference refContract, string logicalName)
        {
            traceService.Trace("CreateRefund");

            Entity newRefund = new Entity("bsd_refund");
            newRefund["bsd_name"] = $"Terminate Refund-{refUnit.Name}";
            newRefund["bsd_customer"] = enTermination.Contains("bsd_customer") ? enTermination["bsd_customer"] : null;
            newRefund["bsd_project"] = enTermination.Contains("bsd_project") ? enTermination["bsd_project"] : null;
            newRefund["bsd_refundtype"] = new OptionSetValue(100000001);    //Terminate Refund
            newRefund["bsd_unitno"] = refUnit;
            newRefund[logicalName] = refContract;
            newRefund["bsd_paymentactualtime"] = DateTime.UtcNow;
            newRefund["bsd_totalamountpaid"] = enTermination.Contains("bsd_totalamountpaid") ? enTermination["bsd_totalamountpaid"] : null;
            newRefund["bsd_refundableamount"] = enTermination.Contains("bsd_totalamountpaid") ? enTermination["bsd_totalamountpaid"] : null;
            newRefund["bsd_source"] = enTermination.Contains("bsd_source") ? enTermination["bsd_source"] : null;

            newRefund.Id = Guid.NewGuid();
            service.Create(newRefund);
        }
    }
}