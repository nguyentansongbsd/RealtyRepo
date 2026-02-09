using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Termination_Approved
{
    public class Plugin_Termination_Approved : IPlugin
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
                Entity enTermination = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "statuscode", "bsd_source", "bsd_reservation",
                    "bsd_reservationcontract", "bsd_optionentry", "bsd_units", "bsd_source", "bsd_resell", "bsd_customer", "bsd_project", "bsd_totalamountpaid", "bsd_forfeitureamount"}));
                int status = enTermination.Contains("statuscode") ? ((OptionSetValue)enTermination["statuscode"]).Value : -99;
                if (status != 100000004)  //Complete
                    return;

                int bsd_source = ((OptionSetValue)enTermination["bsd_source"]).Value;
                if (bsd_source == 100000000 && enTermination.Contains("bsd_reservation"))   //Deposit
                {
                    UpdateStatus("bsd_reservation", enTermination, 667980004);
                }
                else if (bsd_source == 100000001 && enTermination.Contains("bsd_reservationcontract"))   //Reservation Contract
                {
                    UpdateStatus("bsd_reservationcontract", enTermination, 100000004);
                }
                else if (bsd_source == 100000002 && enTermination.Contains("bsd_optionentry"))   //Option Entry
                {
                    UpdateStatus("bsd_optionentry", enTermination, 100000014);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void UpdateStatus(string logicalName, Entity enTermination, int statusContract)
        {
            traceService.Trace("UpdateStatus");

            #region up reservation, reservation contract, oe
            EntityReference refContract = (EntityReference)enTermination[logicalName];
            Entity upContract = new Entity(refContract.LogicalName, refContract.Id);
            upContract["statecode"] = new OptionSetValue(1);    //active
            upContract["statuscode"] = new OptionSetValue(statusContract);  //Terminated
            service.Update(upContract);
            #endregion

            bool bsd_resell = enTermination.Contains("bsd_resell") ? (bool)enTermination["bsd_resell"] : false;

            #region up unit
            EntityReference refUnit = (EntityReference)enTermination["bsd_units"];
            Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
            upUnit["statuscode"] = new OptionSetValue(bsd_resell ? 100000000 : 1);  //Available, Preparing
            service.Update(upUnit);
            #endregion

            decimal bsd_forfeitureamount = enTermination.Contains("bsd_forfeitureamount") ? ((Money)enTermination["bsd_forfeitureamount"]).Value : 0;
            if (bsd_forfeitureamount > 0)
                CreateRefund(enTermination, refUnit, refContract, logicalName);
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