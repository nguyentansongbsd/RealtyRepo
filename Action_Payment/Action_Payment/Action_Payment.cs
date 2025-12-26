using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Payment
{
    public class Action_Payment : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationServiceFactory _serviceFactory;
        private IOrganizationService _service;
        private ITracingService _tracingService;

        EntityReference Target = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            this.Target = this._context.InputParameters["Target"] as EntityReference;
            Entity enPayment = this._service.Retrieve(this.Target.LogicalName, this.Target.Id, new ColumnSet("bsd_units",
                "bsd_project", "bsd_queue", "bsd_amountpay", "bsd_outstandingbalance", "bsd_amountwaspaid"));
            updatePayment();
            updateQueue(enPayment);
        }
        private void updatePayment()
        {
            try
            {
                Entity enUpdatePayment = new Entity(this.Target.LogicalName, this.Target.Id);
                enUpdatePayment["statuscode"] = new OptionSetValue(100000000);
                enUpdatePayment["bsd_confirmeddate"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, this._service);
                enUpdatePayment["bsd_confirmperson"] = new EntityReference("systemuser", this._context.UserId);
                this._service.Update(enUpdatePayment);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void updateQueue(Entity enPayment) 
        {             
            try
            {
                if (!enPayment.Contains("bsd_queue")) return;
                this._tracingService.Trace("Updating Queue Record...");
                decimal amountPay = enPayment.Contains("bsd_amountpay") ? ((Money)enPayment["bsd_amountpay"]).Value : 0;
                decimal amountWasPaid = enPayment.Contains("bsd_amountwaspaid") ? ((Money)enPayment["bsd_amountwaspaid"]).Value : 0;
                decimal totalPaid = amountWasPaid + amountPay;
                decimal queuingFee = getQueuingFee(enPayment);

                Entity enQueue = new Entity(((EntityReference)enPayment["bsd_queue"]).LogicalName, ((EntityReference)enPayment["bsd_queue"]).Id);
                if(totalPaid == queuingFee)
                    enQueue["bsd_collectedqueuingfee"] = true;
                enQueue["bsd_dateorder"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, this._service);
                enQueue["bsd_queuingfeepaid"] = new Money(totalPaid);
                this._service.Update(enQueue);
                this._tracingService.Trace("Queue Record Updated.");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private decimal getQueuingFee(Entity enPayment)
        {
            if (!enPayment.Contains("bsd_queue")) return 0;
            Entity enQueue = this._service.Retrieve(((EntityReference)enPayment["bsd_queue"]).LogicalName, ((EntityReference)enPayment["bsd_queue"]).Id, new ColumnSet("bsd_queuingfee"));
            if(!enQueue.Contains("bsd_queuingfee")) return 0;
            return ((Money)enQueue["bsd_queuingfee"]).Value;
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
