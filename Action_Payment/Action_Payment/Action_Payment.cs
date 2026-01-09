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
            _context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            _serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            _service = _serviceFactory.CreateOrganizationService(_context.UserId);
            _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            Target = _context.InputParameters["Target"] as EntityReference;
            Entity enPayment = getTarget();
            if (!enPayment.Contains("bsd_paymenttype"))
            {
                throw new InvalidPluginExecutionException("No Payment Type value");
            }
            if (!enPayment.Contains("bsd_amountpay"))
            {
                throw new InvalidPluginExecutionException("No Amount Pay value");
            }
            if (!enPayment.Contains("bsd_paymentactualtime"))
            {
                throw new InvalidPluginExecutionException("No Receipt Date value");
            }
            var paymentType = (PaymentType)((OptionSetValue)enPayment["bsd_paymenttype"]).Value;
            decimal amountPay = ((Money)enPayment["bsd_amountpay"]).Value;
            switch (paymentType)
            {
                case PaymentType.QueuingFee:
                    // logic Queuing fee
                    updateQueue(enPayment, amountPay);
                    break;

                case PaymentType.DepositFee:
                    // logic Deposit fee
                    if (!enPayment.Contains("bsd_quotationreservation"))
                    {
                        throw new InvalidPluginExecutionException("No Quotation Reservation value");
                    }
                    deposit(enPayment, (EntityReference)enPayment["bsd_quotationreservation"], amountPay);
                    break;

                case PaymentType.Installment:
                    // logic Installment
                    break;

                case PaymentType.InterestCharge:
                    // logic Interest Charge
                    break;

                case PaymentType.Fees:
                    // logic Fees
                    break;

                case PaymentType.Other:
                    // logic Other
                    break;

                default:
                    throw new InvalidPluginExecutionException("Unsupported Payment Type");
            }
            // cập nhật sts = Paid cho phiếu thu và ghi nhận ngày + người thanh toán
            updatePayment();
        }
        //get target entity Payment
        private Entity getTarget()
        {
            return _service.Retrieve(Target.LogicalName, Target.Id,
                new ColumnSet(
                    "bsd_units",
                    "bsd_project",
                    "bsd_queue",
                    "bsd_quotationreservation",
                    "bsd_reservationcontract",
                    "bsd_optionentry",
                    "bsd_paymenttype",
                    "bsd_transactiontype",
                    "bsd_paymentactualtime",
                    "bsd_amountpay",
                    "bsd_outstandingbalance",
                    "bsd_amountwaspaid"
                )
            );
        }
        // case thanh toán tiền cọc
        private void deposit(Entity enPayment, EntityReference enrDeposit, decimal amountPay)
        {
            Entity enDeposit = _service.Retrieve(enrDeposit.LogicalName, enrDeposit.Id,
                new ColumnSet(
                    "bsd_depositfee",
                    "bsd_totalamountpaid"
                )
            );
            decimal depositfee = enDeposit.Contains("bsd_depositfee") ? ((Money)enDeposit["bsd_depositfee"]).Value : 0;
            decimal totalamountpaid = enDeposit.Contains("bsd_totalamountpaid") ? ((Money)enDeposit["bsd_totalamountpaid"]).Value : 0;
            decimal balance = depositfee - totalamountpaid;
            if (balance <= 0) throw new InvalidPluginExecutionException("The deposit has been paid in full.");
            if (amountPay > balance) throw new InvalidPluginExecutionException("The amount payable is more than the deposit required.");
            Entity upDeposit = new Entity(enrDeposit.LogicalName, enrDeposit.Id);
            totalamountpaid += amountPay;
            upDeposit["bsd_totalamountpaid"] = new Money(totalamountpaid);
            if (totalamountpaid == depositfee) upDeposit["bsd_deposittime"] = enPayment["bsd_paymentactualtime"];
            _service.Update(upDeposit);

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_paymentschemedetail"">
                <attribute name=""bsd_paymentschemedetailid"" />
                <attribute name=""bsd_amountofthisphase"" />
                <attribute name=""bsd_depositamount"" />
                <attribute name=""bsd_amountwaspaid"" />
                <attribute name=""bsd_balance"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""{0}"" />
                  <condition attribute=""bsd_reservation"" operator=""eq"" value=""{enrDeposit.Id}"" />
                  <condition attribute=""bsd_ordernumber"" operator=""eq"" value=""{1}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection enIntallment = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (enIntallment.Entities.Count == 0) throw new InvalidPluginExecutionException("Installment not found.");
            foreach (Entity entity in enIntallment.Entities)
            {
                decimal bsd_amountofthisphase = entity.Contains("bsd_amountofthisphase") ? ((Money)entity["bsd_amountofthisphase"]).Value : 0;
                decimal bsd_depositamount = entity.Contains("bsd_depositamount") ? ((Money)entity["bsd_depositamount"]).Value : 0;
                decimal bsd_amountwaspaid = entity.Contains("bsd_amountwaspaid") ? ((Money)entity["bsd_amountwaspaid"]).Value : 0;
                decimal bsd_balance = entity.Contains("bsd_balance") ? ((Money)entity["bsd_balance"]).Value : 0;
                Entity upIntallment = new Entity(entity.LogicalName, entity.Id);
                upIntallment["bsd_depositamount"] = new Money(bsd_depositamount + amountPay);
                upIntallment["bsd_balance"] = new Money(bsd_amountofthisphase - bsd_depositamount - amountPay - bsd_amountwaspaid);
                _service.Update(upIntallment);
            }
        }
        // cập nhật sts = Paid cho phiếu thu và ghi nhận ngày + người thanh toán
        private void updatePayment()
        {
            try
            {
                Entity enUpdatePayment = new Entity(Target.LogicalName, Target.Id);
                enUpdatePayment["statuscode"] = new OptionSetValue(100000000);
                enUpdatePayment["bsd_confirmeddate"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, _service);
                enUpdatePayment["bsd_confirmperson"] = new EntityReference("systemuser", _context.UserId);
                _service.Update(enUpdatePayment);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        // xử lý thanh toán với case queue
        private void updateQueue(Entity enPayment, decimal amountPay)
        {
            try
            {
                if (!enPayment.Contains("bsd_queue")) return;
                _tracingService.Trace("Updating Queue Record...");
                decimal amountWasPaid = enPayment.Contains("bsd_amountwaspaid") ? ((Money)enPayment["bsd_amountwaspaid"]).Value : 0;
                decimal totalPaid = amountWasPaid + amountPay;
                decimal queuingFee = getQueuingFee(enPayment);

                Entity enQueue = new Entity(((EntityReference)enPayment["bsd_queue"]).LogicalName, ((EntityReference)enPayment["bsd_queue"]).Id);
                if (totalPaid == queuingFee)
                    enQueue["bsd_collectedqueuingfee"] = true;
                enQueue["bsd_dateorder"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, _service);
                enQueue["bsd_queuingfeepaid"] = new Money(totalPaid);
                _service.Update(enQueue);
                _tracingService.Trace("Queue Record Updated.");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private decimal getQueuingFee(Entity enPayment)
        {
            if (!enPayment.Contains("bsd_queue")) return 0;
            Entity enQueue = _service.Retrieve(((EntityReference)enPayment["bsd_queue"]).LogicalName, ((EntityReference)enPayment["bsd_queue"]).Id, new ColumnSet("bsd_queuingfee"));
            if (!enQueue.Contains("bsd_queuingfee")) return 0;
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
        private enum PaymentType
        {
            QueuingFee = 100000000,
            DepositFee = 100000001,
            Installment = 100000002,
            InterestCharge = 100000003,
            Fees = 100000004,
            Other = 100000005
        }
    }
}
