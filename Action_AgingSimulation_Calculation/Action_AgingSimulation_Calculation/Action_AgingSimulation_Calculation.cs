using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text;

namespace Action_AgingSimulation_Calculation
{
    public class Action_AgingSimulation_Calculation : IPlugin
    {
        public IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        public ITracingService traceService = null;
        StringBuilder strMess = new StringBuilder();
        StringBuilder strMess2 = new StringBuilder();
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            string input01 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input01"]))
            {
                input01 = context.InputParameters["input01"].ToString();
            }
            string input02 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input02"]))
            {
                input02 = context.InputParameters["input02"].ToString();
            }
            string input03 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input03"]))
            {
                input03 = context.InputParameters["input03"].ToString();
            }
            string input04 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input04"]))
            {
                input04 = context.InputParameters["input04"].ToString();
            }
            if (input01 == "Buoc 01" && input02 != "")
            {
                traceService.Trace("Bước 01");
                Entity enTarget = new Entity("bsd_interestsimulation");
                enTarget.Id = Guid.Parse(input02);
                enTarget["bsd_powerautomate"] = true;
                enTarget["bsd_calculator"] = true;
                service.Update(enTarget);
            }
            else if (input01 == "Buoc 02" && input02 != "" && input03 != "" && input04 != "")
            {
                traceService.Trace("Bước 02");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                var fetchXml = $@"
                            <fetch>
                              <entity name='bsd_aginginterestsimulationoption'>
                                <all-attributes />
                                <filter type='and'>
                                  <condition attribute='bsd_aginginterestsimulationoptionid' operator='eq' value='{input03}'/>
                                </filter>
                              </entity>
                            </fetch>";
                EntityCollection lstInterestSimulationOption = service.RetrieveMultiple(new FetchExpression(fetchXml));
                foreach (var InterestOption in lstInterestSimulationOption.Entities)
                {
                    CreateAgingDetail(InterestOption);
                }
            }
            else if (input01 == "Buoc 03" && input02 != "" && input04 != "")
            {
                traceService.Trace("Bước 03");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                Entity enConfirmPayment = new Entity("bsd_interestsimulation");
                enConfirmPayment.Id = Guid.Parse(input02);
                enConfirmPayment["bsd_powerautomate"] = false;
                enConfirmPayment["bsd_calculator"] = false;
                enConfirmPayment["bsd_errorincalculation"] = "";
                service.Update(enConfirmPayment);
            }
        }
        private void CreateAgingDetail(Entity InterestOption)
        {
            #region Code tạo detail
            var optinentryid = ((EntityReference)InterestOption["bsd_optionentry"]).Id.ToString();
            var aginginterestsimulationoptionid = InterestOption.Id.ToString();
            var simulationoptions = service.Retrieve("bsd_aginginterestsimulationoption", new Guid(aginginterestsimulationoptionid), new ColumnSet(true));
            Entity enOptionEntry1 = service.Retrieve("bsd_salesorder", new Guid(optinentryid), new ColumnSet(true));
            Entity enInterestSimulation = service.Retrieve("bsd_interestsimulation", ((EntityReference)simulationoptions["bsd_aginginterestsimulation"]).Id, new ColumnSet(true));
            DateTime simulationDate = RetrieveLocalTimeFromUTCTime((DateTime)enInterestSimulation["bsd_simulationdate"]);
            //DELETE RECORDS OLD
            QueryExpression q1 = new QueryExpression("bsd_paymentschemedetail");
            q1.ColumnSet = new ColumnSet(true);
            q1.Criteria = new FilterExpression(LogicalOperator.And);
            q1.Criteria.AddCondition(new ConditionExpression("bsd_optionentry", ConditionOperator.Equal, enOptionEntry1.Id));
            var listInstallment = service.RetrieveMultiple(q1);
            Entity enUint112 = service.Retrieve(((EntityReference)enOptionEntry1["bsd_unitnumber"]).LogicalName, ((EntityReference)enOptionEntry1["bsd_unitnumber"]).Id, new ColumnSet(true));

            if (listInstallment.Entities.Count > 0)
            {
                decimal interestProjectDaily = 0;
                #region CREATE INTEREST SIMULATION DETAIL
                foreach (Entity ins in listInstallment.Entities)
                {
                    //Cập nhật thêm trường thông tin Aging/ Interest Simulation Option khi tạo Aging/ Interest Simulation Detail"
                    createAgingInterestSimulationDetail(enOptionEntry1, enUint112, enInterestSimulation, ins, simulationoptions, simulationDate, interestProjectDaily);
                }
                #endregion
            }
            updateNewInterestAmount(enOptionEntry1, InterestOption);
            updateAdvantPayment(enOptionEntry1, aginginterestsimulationoptionid);
            #endregion
        }
        public DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(service);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            LocalTimeFromUtcTimeResponse response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
            return response.LocalTime;
            //var utcTime = utcTime.ToString("MM/dd/yyyy HH:mm:ss");
            //var localDateOnly = response.LocalTime.ToString("dd-MM-yyyy");
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
        private void createAgingInterestSimulationDetail(Entity oe, Entity UnitsEn, Entity enInterestSimulation, Entity ins, Entity simulationoptions, DateTime simulationDate, decimal interestProjectDaily)
        {
            // GET INTEREST CHARGE MASTER
            if (!oe.Contains("bsd_paymentscheme"))
                throw new InvalidPluginExecutionException("Please input Payment Scheme in SPA: " + (oe.Contains("bsd_name") ? oe["bsd_name"] : ""));

            Entity payScheme = service.Retrieve(((EntityReference)oe["bsd_paymentscheme"]).LogicalName, ((EntityReference)oe["bsd_paymentscheme"]).Id,
                new ColumnSet(new string[2] {
                                "bsd_name",
                                "bsd_interestratemaster"
                }));

            if (!payScheme.Contains("bsd_interestratemaster"))
                throw new InvalidPluginExecutionException("Please input Interest Charge Master in Payment Scheme: " + (payScheme.Contains("bsd_name") ? payScheme["bsd_name"] : ""));
            DateTime InterestStarDate = new DateTime();
            int lateDays = 0;
            decimal interest_New = 0;
            decimal interest = 0;
            decimal decInterestCharge = 0;
            decimal bsd_interestpercent = 0;
            int bsd_numberofdaysdue;
            // Số tiền trễ chưa thanh toán
            decimal interest_NotPaid = 0;
            interest_NotPaid = CalculateInterestNotPaid(ins);
            var decInterestAmountIns = CalculateInterestAmount(ins);
            Installment objIns = new Installment();
            //Gọi tính trễ

            if (ins.Contains("bsd_duedate"))
            {
                int bsd_gracedays = ins.Contains("bsd_gracedays") ? (int)ins["bsd_gracedays"] : 0;
                DateTime dueDate = RetrieveLocalTimeFromUTCTime((DateTime)ins["bsd_duedate"]).AddDays(bsd_gracedays);
                InterestStarDate = dueDate.AddDays(1);
                dueDate = new DateTime(dueDate.Year, dueDate.Month, dueDate.Day);
                simulationDate = new DateTime(simulationDate.Year, simulationDate.Month, simulationDate.Day);
                bool bsd_official = ins.Contains("bsd_official") ? (bool)ins["bsd_official"] : false;
                if (bsd_official)
                {
                    int gracedays = (int)simulationDate.Date.Subtract(dueDate.Date).TotalDays;
                    if (gracedays > 0)
                    {
                        int Gracedays = gracedays < 0 ? 0 : gracedays;
                        bsd_interestpercent = ins.Contains("bsd_interestpercent") ? (decimal)ins["bsd_interestpercent"] : 0;
                        interest_New = Math.Round(interest_NotPaid * gracedays * bsd_interestpercent / 100, MidpointRounding.AwayFromZero);
                    }
                }
            }
            string UnitsID = UnitsEn.Id.ToString();
            string UnitsName = UnitsEn.Contains("bsd_name") ? UnitsEn["bsd_name"].ToString() : "";

            Entity ISDetail = new Entity("bsd_interestsimulationdetail");
            if (enInterestSimulation != null)
            {
                ISDetail["bsd_name"] = "" + (enInterestSimulation.Contains("bsd_name") ? enInterestSimulation["bsd_name"] : "") + "-" + UnitsName + "-" + (ins.Contains("bsd_name") ? ins["bsd_name"] : "");
                ISDetail["bsd_interestsimulation"] = enInterestSimulation.ToEntityReference();
            }
            else
            {
                ISDetail["bsd_name"] = UnitsName + "-" + (ins.Contains("bsd_name") ? ins["bsd_name"] : "");
                ISDetail["bsd_interestsimulation"] = null;
            }
            ISDetail["bsd_optionentry"] = oe.ToEntityReference();
            ISDetail["bsd_simulationdate"] = simulationDate;
            traceService.Trace("simulationDate " + simulationDate);
            ISDetail["bsd_type"] = new OptionSetValue(100000002);
            ISDetail["bsd_paymentscheme"] = payScheme.ToEntityReference();
            ISDetail["bsd_units"] = UnitsEn.ToEntityReference();
            ISDetail["bsd_aginginterestsimulationoption"] = simulationoptions != null ? simulationoptions.ToEntityReference() : null;
            ISDetail["bsd_installment"] = ins.ToEntityReference();
            ISDetail["bsd_installmentamount"] = ins.Contains("bsd_amountofthisphase") ? ins["bsd_amountofthisphase"] : new Money(0);
            ISDetail["bsd_paidamount"] = ins.Contains("bsd_amountwaspaid") ? ins["bsd_amountwaspaid"] : new Money(0);
            ISDetail["bsd_outstandingamount"] = ins.Contains("bsd_balance") ? ins["bsd_balance"] : new Money(0);
            ISDetail["bsd_numberofdaysdue"] = 0;
            // Create Detail
            interest = interest_New + interest_NotPaid;
            #region -- type = Interest Simulation = 100000001
            DateTime duedate = new DateTime();
            decimal bsd_interestchargeamount = 0;
            bool bolCheckPaid = false;
            bolCheckPaid = (((OptionSetValue)ins["statuscode"]).Value == 100000001 && ins.Contains("bsd_interestchargestatus") && ((OptionSetValue)ins["bsd_interestchargestatus"]).Value == 100000000);
            if (interest_NotPaid > 0 || ((OptionSetValue)ins["statuscode"]).Value == 100000000 || bolCheckPaid)
            {
                if (ins.Contains("bsd_duedate"))
                {
                    duedate = RetrieveLocalTimeFromUTCTime((DateTime)ins["bsd_duedate"]);
                    ISDetail["bsd_duedate"] = duedate;
                    DateTime InterestStarDate1 = InterestStarDate;
                    ISDetail["bsd_intereststartdate"] = InterestStarDate1;
                    traceService.Trace("InterestStarDate1 " + InterestStarDate1);
                    bsd_numberofdaysdue = (int)simulationDate.Date.Subtract(duedate.Date).TotalDays;
                    if (bsd_numberofdaysdue > 0)
                    {
                        ISDetail["bsd_numberofdaysdue"] = bsd_numberofdaysdue;
                    }
                    traceService.Trace("gán lateDays: " + lateDays);
                    ISDetail["bsd_outstandingday"] = lateDays;
                    ISDetail["bsd_paymentscheme"] = payScheme.ToEntityReference();
                    ISDetail["bsd_interestpercent"] = bsd_interestpercent;
                    ISDetail["bsd_groupaging"] = new OptionSetValue(CheckGroupAging(lateDays));
                    ISDetail["bsd_interestamountinstallment"] = new Money(decInterestAmountIns);
                    decimal decnewinterestamount = bolCheckPaid ? 0 : interest_New > decInterestCharge ? decInterestCharge : interest_New;
                    ISDetail["bsd_newinterestamount"] = new Money(decnewinterestamount);
                    ISDetail["bsd_interestchargeamount"] = new Money(decnewinterestamount + decInterestAmountIns);
                    bsd_interestchargeamount = interest + interest_New;
                    ISDetail["bsd_advancepayment"] = new Money(0);
                }
                else
                {
                    ISDetail["bsd_outstandingday"] = 0;
                    ISDetail["bsd_paymentscheme"] = payScheme.ToEntityReference();
                    ISDetail["bsd_interestpercent"] = (decimal)0;
                    ISDetail["bsd_interestamountinstallment"] = new Money(0);
                    ISDetail["bsd_newinterestamount"] = new Money(0);
                    ISDetail["bsd_interestchargeamount"] = new Money(0);
                    ISDetail["bsd_advancepayment"] = new Money(0);
                }

            }
            #endregion
            else
            {
                if (ins.Contains("bsd_duedate"))
                {
                    duedate = RetrieveLocalTimeFromUTCTime((DateTime)ins["bsd_duedate"]);
                    ISDetail["bsd_duedate"] = duedate;
                    ISDetail["bsd_outstandingday"] = 0;
                    ISDetail["bsd_paymentscheme"] = payScheme.ToEntityReference();
                    ISDetail["bsd_interestpercent"] = (decimal)0;
                    ISDetail["bsd_interestamountinstallment"] = new Money(0);
                    ISDetail["bsd_newinterestamount"] = new Money(0);
                    ISDetail["bsd_interestchargeamount"] = new Money(0);
                    ISDetail["bsd_advancepayment"] = new Money(0);
                }
            }
            service.Create(ISDetail);
            traceService.Trace("qua Create(ISDetail) ");
        }
        private EntityCollection getOptionEntrys(IOrganizationService crmservices, string idUnit)
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendLine("<fetch version='1.0' output-format='xml-platform' mapping='logical'>");
            xml.AppendLine("<entity name='bsd_salesorder'>");
            xml.AppendLine("<attribute name='bsd_name' />");
            xml.AppendLine("<attribute name='bsd_totalamount' />");
            xml.AppendLine("<attribute name='statuscode' />");
            xml.AppendLine("<attribute name='bsd_customerid' />");
            xml.AppendLine("<attribute name='createdon' />");
            xml.AppendLine("<attribute name='bsd_unitnumber' />");
            xml.AppendLine("<attribute name='bsd_project' />");
            xml.AppendLine("<attribute name='bsd_optionno' />");
            xml.AppendLine("<attribute name='bsd_contractnumber' />");
            xml.AppendLine("<attribute name='bsd_optioncodesams' />");
            xml.AppendLine("<attribute name='bsd_salesorderid' />"); ;
            xml.AppendLine("<filter type='and'>");
            xml.AppendLine(string.Format("<condition attribute='bsd_unitnumber' operator='eq' value='{0}'/>", idUnit));
            xml.AppendLine("<condition attribute='statuscode' operator='ne' value='100000014'/>");
            xml.AppendLine("<condition attribute='statuscode' operator='ne' value='100000012'/>");
            xml.AppendLine("<condition attribute='statuscode' operator='ne' value='2'/>");
            xml.AppendLine("</filter>");
            xml.AppendLine("</entity>");
            xml.AppendLine("</fetch>");
            traceService.Trace(xml.ToString());
            EntityCollection entc = service.RetrieveMultiple(new FetchExpression(xml.ToString()));
            return entc;
        }
        private decimal getInterestCap(Entity enOptionEntry)
        {
            decimal lim = -100599;
            decimal totalamount = enOptionEntry.Contains("bsd_totalamount") ? ((Money)enOptionEntry["bsd_totalamount"]).Value : 0;
            EntityReference enrefPaymentScheme = enOptionEntry.Contains("bsd_paymentscheme") ? (EntityReference)enOptionEntry["bsd_paymentscheme"] : null;
            if (enrefPaymentScheme != null)
            {
                Entity enPaymentScheme = service.Retrieve(enrefPaymentScheme.LogicalName, enrefPaymentScheme.Id, new ColumnSet(true));
                EntityReference enrefInterestRateMaster = enPaymentScheme.Contains("bsd_interestratemaster") ? (EntityReference)enPaymentScheme["bsd_interestratemaster"] : null;
                if (enrefInterestRateMaster != null)
                {
                    Entity enInterestRateMaster = service.Retrieve(enrefInterestRateMaster.LogicalName, enrefInterestRateMaster.Id, new ColumnSet(true));
                    decimal bsd_toleranceinterestamount = enInterestRateMaster.Contains("bsd_toleranceinterestamount") ? ((Money)enInterestRateMaster["bsd_toleranceinterestamount"]).Value : 0;
                    decimal bsd_toleranceinterestpercentage = enInterestRateMaster.Contains("bsd_toleranceinterestpercentage") ? (decimal)enInterestRateMaster["bsd_toleranceinterestpercentage"] : 0;
                    decimal amountcalbypercent = totalamount * bsd_toleranceinterestpercentage / 100;
                    if (bsd_toleranceinterestamount > 0 && amountcalbypercent > 0 && enInterestRateMaster.Contains("bsd_toleranceinterestamount") && enInterestRateMaster.Contains("bsd_toleranceinterestpercentage"))
                    {
                        lim = Math.Min(bsd_toleranceinterestamount, amountcalbypercent);
                    }
                    else
                    {
                        if (bsd_toleranceinterestamount > 0 && enInterestRateMaster.Contains("bsd_toleranceinterestamount"))
                        {
                            lim = bsd_toleranceinterestamount;
                        }
                        if (amountcalbypercent > 0 && enInterestRateMaster.Contains("bsd_toleranceinterestpercentage"))
                        {
                            lim = amountcalbypercent;
                        }
                    }

                }
            }
            return lim;
        }
        private void updateNewInterestAmount(Entity enOptionEntry, Entity enInterestSimulateOption)
        {
            Entity optionEntry = service.Retrieve(enOptionEntry.LogicalName, enOptionEntry.Id, new ColumnSet(true));

            //Lấy list Interest Simulation Detail có New Interest Amount > 0 theo thứ tự
            string condition = "<condition attribute='bsd_aginginterestsimulationoption' operator='eq' uitype='bsd_aginginterestsimulationoption' value='" + enInterestSimulateOption.Id.ToString() + "' />";
            string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
	            <entity name='bsd_interestsimulationdetail'>
		            <attribute name='bsd_interestsimulationdetailid' />
		            <attribute name='bsd_name' />
		            <attribute name='createdon' />
		            <attribute name='bsd_interestamountinstallment' />
		            <attribute name='bsd_interestchargeamount' />
		            <attribute name='bsd_newinterestamount' />
		            <attribute name='bsd_interestsimulation' />
		            <attribute name='bsd_inrerestamountinstallment' />
		            <filter type='and'>
			            <condition attribute='bsd_optionentry' operator='eq' uitype='salesorder' value='" + enOptionEntry.Id.ToString() + @"' />
			            " + condition + @"
		            </filter>
		            <link-entity name='bsd_paymentschemedetail' from='bsd_paymentschemedetailid' to='bsd_installment' visible='false' link-type='outer' alias='interestsimulationdetail_paymentschemedetail'>
			            <attribute name='bsd_ordernumber' />
			            <order attribute='bsd_ordernumber' descending='false' />
		            </link-entity>
	            </entity>
            </fetch>";
            EntityCollection encolInterestSimulationDetail = service.RetrieveMultiple(new FetchExpression(xml));
            decimal cap = getInterestCap(optionEntry);
            if (cap != -100599)
            {
                //Tính toán New Interest Amount sao cho tổng nhỏ hơn hoặc bằng cap
                //Ưu tiên giảm các đợt cuối trở về trước
                decimal sumInterestAmount = encolInterestSimulationDetail.Entities.AsEnumerable().Sum(x => x.Contains("bsd_interestamountinstallment") ? ((Money)x["bsd_interestamountinstallment"]).Value : 0);
                decimal[] arrNewInterestAmount = { };
                for (int i = 0; i < encolInterestSimulationDetail.Entities.Count; i++)
                {
                    Entity enInterestSimulationDetail = encolInterestSimulationDetail.Entities[i];
                    decimal bsd_newinterestamount = enInterestSimulationDetail.Contains("bsd_newinterestamount") ? ((Money)enInterestSimulationDetail["bsd_newinterestamount"]).Value : 0;
                    decimal bsd_interestamountinstallment = enInterestSimulationDetail.Contains("bsd_interestamountinstallment") ? ((Money)enInterestSimulationDetail["bsd_interestamountinstallment"]).Value : 0;
                    if (sumInterestAmount < cap)
                    {
                        decimal total = sumInterestAmount + bsd_newinterestamount;
                        if (total > cap && bsd_newinterestamount != 0)
                        {
                            decimal denta = cap - sumInterestAmount;
                            //Set New Interest Amount = denta
                            Entity en = new Entity(enInterestSimulationDetail.LogicalName, enInterestSimulationDetail.Id);
                            en["bsd_newinterestamount"] = new Money(denta);
                            en["bsd_interestchargeamount"] = new Money(denta + bsd_interestamountinstallment);
                            service.Update(en);
                            sumInterestAmount += denta;
                            traceService.Trace("vào: " + 1);
                        }
                        else
                        {
                            sumInterestAmount += bsd_newinterestamount;
                            traceService.Trace("vào: " + 2);
                        }
                    }
                    else
                    {
                        //Set New Interest Amount = 0
                        Entity en = new Entity(enInterestSimulationDetail.LogicalName, enInterestSimulationDetail.Id);
                        en["bsd_newinterestamount"] = new Money(0);
                        en["bsd_interestchargeamount"] = new Money(bsd_interestamountinstallment);
                        service.Update(en);
                        traceService.Trace("vào: " + 3);
                    }
                }
            }
        }
        private decimal CalculateInterestNotPaid(Entity ins)
        {
            decimal interestamount = ins.Contains("bsd_amountofthisphase") ? ((Money)ins["bsd_amountofthisphase"]).Value : decimal.Zero;
            decimal interestamountpaid = ins.Contains("bsd_amountwaspaid") ? ((Money)ins["bsd_amountwaspaid"]).Value : decimal.Zero;
            decimal waiverinterest = ins.Contains("bsd_waiverinstallment") ? ((Money)ins["bsd_waiverinstallment"]).Value : decimal.Zero;
            return interestamount - interestamountpaid - waiverinterest;
        }
        private decimal CalculateInterestAmount(Entity ins)
        {
            decimal interestamount = ins.Contains("bsd_interestchargeamount") ? ((Money)ins["bsd_interestchargeamount"]).Value : decimal.Zero;
            decimal interestamountpaid = ins.Contains("bsd_interestwaspaid") ? ((Money)ins["bsd_interestwaspaid"]).Value : decimal.Zero;
            decimal waiverinterest = ins.Contains("bsd_waiverinterest") ? ((Money)ins["bsd_waiverinterest"]).Value : decimal.Zero;
            return interestamount - interestamountpaid - waiverinterest;
        }
        private int CheckGroupAging(decimal i)
        {
            if (i <= 15)
                return 100000000;
            else if (i > 15 && i <= 30)
                return 100000001;
            else if (i > 30 && i <= 60)
                return 100000002;
            else if (i > 60 && i <= 90)
                return 100000003;
            else // (i > 90)
                return 100000004;
        }
        private EntityCollection CalSum_AdvancePayment(string OptionID)
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendLine("<fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>");
            xml.AppendLine("<entity name='bsd_advancepayment'>");
            xml.AppendLine("<attribute name='bsd_remainingamount' alias='SumAdv' aggregate='sum'/>");
            xml.AppendLine("<filter type='and'>");
            xml.AppendLine(string.Format("<condition attribute='bsd_optionentry' operator='eq' value='{0}'/>", OptionID));
            xml.AppendLine("<condition attribute='statuscode' operator='eq' value='100000000'/>");
            xml.AppendLine("</filter>");
            xml.AppendLine("</entity>");
            xml.AppendLine("</fetch>");
            return service.RetrieveMultiple(new FetchExpression(xml.ToString()));
        }
        private void updateAdvantPayment(Entity oe, string aginginterestsimulationoption)
        {
            //Cal Sum Advance Payment
            decimal AdvPayAmt = 0;
            EntityCollection AdvSum = CalSum_AdvancePayment(oe.Id.ToString());
            if (AdvSum.Entities.Count > 0)
            {
                Entity AdvSumEn = AdvSum.Entities[0];
                if (((AliasedValue)AdvSumEn.Attributes["SumAdv"]).Value != null)
                    AdvPayAmt = ((Money)((AliasedValue)AdvSumEn.Attributes["SumAdv"]).Value).Value;
            }
            string condition = "<condition attribute='bsd_aginginterestsimulationoption' operator='null' />";
            if (aginginterestsimulationoption != "")
            {
                condition = "<condition attribute='bsd_aginginterestsimulationoption' operator='eq' value='" + aginginterestsimulationoption + "'/>";
            }
            string xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
	            <entity name='bsd_interestsimulationdetail'>
		            <attribute name='bsd_name' />
		            <attribute name='createdon' />
		            <attribute name='bsd_units' />
		            <attribute name='bsd_optionentry' />
		            <attribute name='bsd_installment' />
		            <attribute name='statuscode' />
		            <attribute name='bsd_outstandingday' />
		            <attribute name='bsd_interestsimulation' />
		            <attribute name='bsd_interestchargeamount' />
		            <attribute name='bsd_interestpercent' />
		            <attribute name='bsd_aginginterestsimulationoption' />
		            <attribute name='bsd_interestsimulationdetailid' />
		            <attribute name='bsd_duedate' />
		            <attribute name='bsd_simulationdate' />
		            <attribute name='bsd_paidamount' />
		            <attribute name='bsd_installmentamount' />
		            <attribute name='bsd_outstandingamount' />
                    
		            <filter type='and'>
			            <condition attribute='statecode' operator='eq' value='0' />
			            <condition attribute='bsd_optionentry' operator='eq' uitype='salesorder' value='" + oe.Id.ToString() + @"' />
			            " + condition + @"
                       
		            </filter>
		            <link-entity name='bsd_paymentschemedetail' from='bsd_paymentschemedetailid' to='bsd_installment' visible='false' link-type='outer' alias='installment'>
			            <attribute name='bsd_ordernumber' />
			            <order attribute='bsd_ordernumber' descending='false' />
		            </link-entity>
	            </entity>
            </fetch>";
            //Hồ fix 03-06-2019
            EntityCollection encoInterestSimulationDetail = service.RetrieveMultiple(new FetchExpression(xml));
            if (encoInterestSimulationDetail.Entities.Count > 0)
            {
                Entity enInterestSimulationDetail = new Entity("bsd_interestsimulationdetail", encoInterestSimulationDetail.Entities[0].Id);
                enInterestSimulationDetail["bsd_advancepayment"] = new Money(AdvPayAmt);
                service.Update(enInterestSimulationDetail);
            }
        }
    }
    public class Installment
    {
        public DateTime InterestStarDate { get; set; }
        public int Intereststartdatetype { get; set; }
        public int Gracedays { get; set; }
        public int LateDays { get; set; }
        public int orderNumber { get; set; }
        public Guid idOE { get; set; }
        public decimal MaxPercent { get; set; }
        public decimal MaxAmount { get; set; }
        public decimal InterestPercent { get; set; }
        public decimal InterestCharge { get; set; }
        public DateTime Duedate { get; set; }
    }
}