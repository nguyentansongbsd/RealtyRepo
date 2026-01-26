using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
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
                    if (!enPayment.Contains("bsd_transactiontype"))
                    {
                        throw new InvalidPluginExecutionException("No Transaction Type value");
                    }
                    int bsd_transactiontype = ((OptionSetValue)enPayment["bsd_transactiontype"]).Value;
                    if (bsd_transactiontype == 100000000 && !enPayment.Contains("bsd_reservationcontract"))
                        throw new InvalidPluginExecutionException("No Reservation Contract value");
                    else if (bsd_transactiontype == 100000001 && !enPayment.Contains("bsd_optionentry"))
                        throw new InvalidPluginExecutionException("No Optionentry value");
                    installment(enPayment, (bsd_transactiontype == 100000000 ? (EntityReference)enPayment["bsd_reservationcontract"] : (EntityReference)enPayment["bsd_optionentry"]), amountPay);
                    break;

                case PaymentType.InterestCharge:
                    // logic Interest Charge
                    if (!enPayment.Contains("bsd_transactiontype"))
                    {
                        throw new InvalidPluginExecutionException("No Transaction Type value");
                    }
                    int transactiontype = ((OptionSetValue)enPayment["bsd_transactiontype"]).Value;
                    if (transactiontype == 100000000 && !enPayment.Contains("bsd_reservationcontract"))
                        throw new InvalidPluginExecutionException("No Reservation Contract value");
                    else if (transactiontype == 100000001 && !enPayment.Contains("bsd_optionentry"))
                        throw new InvalidPluginExecutionException("No Optionentry value");
                    Interest_Charge(enPayment, (transactiontype == 100000000 ? (EntityReference)enPayment["bsd_reservationcontract"] : (EntityReference)enPayment["bsd_optionentry"]), amountPay);
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
        // case thanh toán tiền lãi
        private void Interest_Charge(Entity enPayment, EntityReference enrHD, decimal amountPay)
        {
            Entity enHD = _service.Retrieve(enrHD.LogicalName, enrHD.Id,
                new ColumnSet(
                    "bsd_totalinterest",
                    "bsd_totalinterestpaid"
                )
            );
            decimal totalinterest = enHD.Contains("bsd_totalinterest") ? ((Money)enHD["bsd_totalinterest"]).Value : 0;
            decimal totalinterestpaid = enHD.Contains("bsd_totalinterestpaid") ? ((Money)enHD["bsd_totalinterestpaid"]).Value : 0;
            decimal balance = totalinterest - totalinterestpaid;
            if (balance <= 0) throw new InvalidPluginExecutionException("The interest charge has been paid in full.");
            if (amountPay > balance) throw new InvalidPluginExecutionException("The amount payable is more than the interest charge required.");
            Entity upHD = new Entity(enrHD.LogicalName, enrHD.Id);
            upHD["bsd_totalinterest"] = new Money(totalinterest);
            upHD["bsd_totalinterestpaid"] = new Money(totalinterestpaid + amountPay);
            upHD["bsd_totalinterestremaining"] = new Money(totalinterest - totalinterestpaid - amountPay);
            _service.Update(upHD);
            string nameField = enrHD.LogicalName == "bsd_reservationcontract" ? "bsd_reservationcontract" : "bsd_optionentry";
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <attribute name=""bsd_paymentschemedetailid"" />
                <attribute name=""bsd_interestchargestatus"" />
                <attribute name=""bsd_interestchargeamount"" />
                <attribute name=""bsd_interestwaspaid"" />
                <attribute name=""bsd_waiverinterest"" />
                <attribute name=""bsd_balance"" />
                <filter>
                  <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                  <condition attribute=""bsd_interestchargeremaining"" operator=""gt"" value=""{0}"" />
                  <condition attribute=""{nameField}"" operator=""eq"" value=""{enrHD.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection enIntallment = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (enIntallment.Entities.Count == 0) throw new InvalidPluginExecutionException("Installment not found.");
            //_tracingService.Trace("enIntallment " + enIntallment.Entities.Count);
            foreach (Entity entity in enIntallment.Entities)
            {
                decimal bsd_balanceIns = entity.Contains("bsd_balance") ? ((Money)entity["bsd_balance"]).Value : 0;
                decimal bsd_interestchargeamount = entity.Contains("bsd_interestchargeamount") ? ((Money)entity["bsd_interestchargeamount"]).Value : 0;
                decimal bsd_interestwaspaid = entity.Contains("bsd_interestwaspaid") ? ((Money)entity["bsd_interestwaspaid"]).Value : 0;
                decimal bsd_waiverinterest = entity.Contains("bsd_waiverinterest") ? ((Money)entity["bsd_waiverinterest"]).Value : 0;
                decimal bsd_balance = bsd_interestchargeamount - bsd_interestwaspaid - bsd_waiverinterest;
                Entity upIntallment = new Entity(entity.LogicalName, entity.Id);
                if (amountPay <= bsd_balance)
                {
                    upIntallment["bsd_interestwaspaid"] = new Money(bsd_interestwaspaid + amountPay);
                    upIntallment["bsd_interestchargeremaining"] = new Money(bsd_balance - amountPay);

                    if (amountPay == bsd_balance && bsd_balanceIns == 0)
                    {
                        upIntallment["statuscode"] = new OptionSetValue(100000001);
                        upIntallment["bsd_interestchargestatus"] = new OptionSetValue(100000001);
                    }
                    amountPay = 0;
                }
                else
                {
                    upIntallment["bsd_interestwaspaid"] = new Money(bsd_interestwaspaid + bsd_balance);
                    upIntallment["bsd_interestchargeremaining"] = new Money(0);
                    if (bsd_balanceIns == 0)
                    {
                        upIntallment["statuscode"] = new OptionSetValue(100000001);
                        upIntallment["bsd_interestchargestatus"] = new OptionSetValue(100000001);
                    }
                    amountPay -= bsd_balance;
                }
                _service.Update(upIntallment);
                if (amountPay <= 0) break;
            }
            if (amountPay > 0)
            {
                throw new InvalidPluginExecutionException("The amount payable is more than the interest charge required.");
            }
        }
        // case thanh toán tiền đợt
        private void installment(Entity enPayment, EntityReference enrHD, decimal amountPay)
        {
            Entity enHD = _service.Retrieve(enrHD.LogicalName, enrHD.Id,
                new ColumnSet(
                    "bsd_totalamount",
                    "bsd_totalinterest",
                    "bsd_totalinterestpaid",
                    "bsd_totalamountpaid"
                )
            );
            decimal totalamount = enHD.Contains("bsd_totalamount") ? ((Money)enHD["bsd_totalamount"]).Value : 0;
            decimal totalamountpaid = enHD.Contains("bsd_totalamountpaid") ? ((Money)enHD["bsd_totalamountpaid"]).Value : 0;
            totalamountpaid += amountPay;
            decimal balance = totalamount - totalamountpaid;
            decimal bsd_totalinterest = enHD.Contains("bsd_totalinterest") ? ((Money)enHD["bsd_totalinterest"]).Value : 0;
            decimal totalinterestpaid = enHD.Contains("bsd_totalinterestpaid") ? ((Money)enHD["bsd_totalinterestpaid"]).Value : 0;
            //_tracingService.Trace("totalamount " + totalamount);
            //_tracingService.Trace("totalamountpaid " + totalamountpaid);
            //_tracingService.Trace("balance " + balance);
            //_tracingService.Trace("amountPay " + amountPay);
            if (balance <= 0) throw new InvalidPluginExecutionException("The installment has been paid in full.");
            if (amountPay > balance) throw new InvalidPluginExecutionException("The amount payable is more than the installment required.");
            string nameField = enrHD.LogicalName == "bsd_reservationcontract" ? "bsd_reservationcontract" : "bsd_optionentry";
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetail"">
                <attribute name=""bsd_paymentschemedetailid"" />
                <attribute name=""bsd_amountofthisphase"" />
                <attribute name=""bsd_depositamount"" />
                <attribute name=""bsd_amountwaspaid"" />
                <attribute name=""bsd_balance"" />
                <attribute name=""bsd_intereststartdate"" />
                <attribute name=""bsd_interestchargestatus"" />
                <attribute name=""bsd_interestchargeamount"" />
                <attribute name=""bsd_interestwaspaid"" />
                <attribute name=""bsd_waiverinterest"" />
                <attribute name=""bsd_gracedays"" />
                <attribute name=""bsd_interestpercent"" />
                <attribute name=""bsd_official"" />
                <attribute name=""bsd_duedate"" />
                <attribute name=""bsd_ordernumber"" />
                <filter>
                  <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                  <condition attribute=""bsd_balance"" operator=""gt"" value=""{0}"" />
                  <condition attribute=""{nameField}"" operator=""eq"" value=""{enrHD.Id}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection enIntallment = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (enIntallment.Entities.Count == 0) throw new InvalidPluginExecutionException("Installment not found.");
            //_tracingService.Trace("enIntallment " + enIntallment.Entities.Count);
            foreach (Entity entity in enIntallment.Entities)
            {
                decimal bsd_amountofthisphase = entity.Contains("bsd_amountofthisphase") ? ((Money)entity["bsd_amountofthisphase"]).Value : 0;
                decimal bsd_depositamount = entity.Contains("bsd_depositamount") ? ((Money)entity["bsd_depositamount"]).Value : 0;
                decimal bsd_amountwaspaid = entity.Contains("bsd_amountwaspaid") ? ((Money)entity["bsd_amountwaspaid"]).Value : 0;
                decimal bsd_balance = bsd_amountofthisphase - bsd_depositamount - bsd_amountwaspaid;
                InterestCharge interest = new InterestCharge();
                caculateLai(entity, enPayment, bsd_balance, interest);
                Entity upIntallment = new Entity(entity.LogicalName, entity.Id);
                int bsd_interestchargestatus = entity.Contains("bsd_interestchargestatus") ? ((OptionSetValue)entity["bsd_interestchargestatus"]).Value : 0;
                decimal bsd_interestchargeamount = entity.Contains("bsd_interestchargeamount") ? ((Money)entity["bsd_interestchargeamount"]).Value : 0;
                decimal bsd_interestwaspaid = entity.Contains("bsd_interestwaspaid") ? ((Money)entity["bsd_interestwaspaid"]).Value : 0;
                if (amountPay <= bsd_balance)
                {
                    upIntallment["bsd_amountwaspaid"] = new Money(bsd_amountwaspaid + amountPay);
                    upIntallment["bsd_balance"] = new Money(bsd_balance - amountPay);

                    if (amountPay == bsd_balance && ((bsd_interestchargestatus == 100000001 && bsd_interestchargeamount > 0) || (bsd_interestchargestatus == 100000000 && bsd_interestchargeamount <= 0) || !interest.isLai))
                        upIntallment["statuscode"] = new OptionSetValue(100000001);
                    amountPay = 0;
                }
                else
                {
                    upIntallment["bsd_amountwaspaid"] = new Money(bsd_amountwaspaid + bsd_balance);
                    upIntallment["bsd_balance"] = new Money(0);
                    if ((bsd_interestchargestatus == 100000001 && bsd_interestchargeamount > 0) || (bsd_interestchargestatus == 100000000 && bsd_interestchargeamount <= 0) && !interest.isLai)
                        upIntallment["statuscode"] = new OptionSetValue(100000001);
                    amountPay -= bsd_balance;
                }
                if (interest.isLai)
                {
                    if (!entity.Contains("bsd_intereststartdate")) upIntallment["bsd_intereststartdate"] = enPayment["bsd_paymentactualtime"];
                    upIntallment["bsd_interestchargeamount"] = new Money(bsd_interestchargeamount + interest.InterestChargeAmount);
                    upIntallment["bsd_interestchargeremaining"] = new Money(bsd_interestchargeamount + interest.InterestChargeAmount - bsd_interestwaspaid);
                    bsd_totalinterest += interest.InterestChargeAmount;
                    upIntallment["bsd_actualgracedays"] = interest.Gracedays;
                }
                _service.Update(upIntallment);
                if (amountPay <= 0) break;
            }
            Entity upHD = new Entity(enrHD.LogicalName, enrHD.Id);

            upHD["bsd_totalinterest"] = new Money(bsd_totalinterest);
            upHD["bsd_totalinterestremaining"] = new Money(bsd_totalinterest - totalinterestpaid);
            upHD["bsd_totalamountpaid"] = new Money(totalamountpaid);
            decimal percenPaid = Math.Round((totalamountpaid / totalamount * 100), 2, MidpointRounding.AwayFromZero);
            upHD["bsd_totalpercent"] = percenPaid;
            _service.Update(upHD);
            if (amountPay > 0)
            {
                throw new InvalidPluginExecutionException("The amount payable is more than the installment required.");
            }
        }
        // tính phát sinh lãi
        private void caculateLai(Entity enInstallment, Entity enPayment, decimal bsd_balance, InterestCharge interest)
        {
            bool bsd_official = enInstallment.Contains("bsd_official") ? (bool)enInstallment["bsd_official"] : false;
            if (bsd_official && enInstallment.Contains("bsd_duedate") && enPayment.Contains("bsd_paymentactualtime"))
            {
                int bsd_gracedays = enInstallment.Contains("bsd_gracedays") ? (int)enInstallment["bsd_gracedays"] : 0;
                DateTime dueDate = RetrieveLocalTimeFromUTCTime((DateTime)enInstallment["bsd_duedate"], _service).AddDays(bsd_gracedays);
                dueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
                DateTime receiptDate = RetrieveLocalTimeFromUTCTime((DateTime)enPayment["bsd_paymentactualtime"], _service);
                receiptDate = new DateTime(receiptDate.Year, receiptDate.Month, receiptDate.Day);
                int gracedays = (int)receiptDate.Date.Subtract(dueDate.Date).TotalDays;
                if (gracedays > 0)
                {
                    interest.Gracedays = gracedays < 0 ? 0 : gracedays;
                    decimal bsd_interestpercent = enInstallment.Contains("bsd_interestpercent") ? (decimal)enInstallment["bsd_interestpercent"] : 0;
                    interest.InterestChargeAmount = Math.Round(bsd_balance * gracedays * bsd_interestpercent / 100, MidpointRounding.AwayFromZero);
                    interest.isLai = true;
                }
            }
            else interest.isLai = false;
        }
        public class InterestCharge
        {
            public bool isLai { get; set; }
            public int Gracedays { get; set; }
            public decimal InterestChargeAmount { get; set; }
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
