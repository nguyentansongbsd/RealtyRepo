using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using RealtyCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_ReservationContract_PSGen
{
    public class Action_ReservationContract_PSGen : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        decimal bsd_freightamount = 0;
        decimal bsd_managementfee = 0;
        ITracingService traceS = null;
        List<DateTime> listCalendar;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            try
            {
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                EntityReference target = (EntityReference)context.InputParameters["Target"];
                traceS = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceS.Trace($"start {target.Id}");

                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                Entity enHD = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));

                traceS.Trace("2");

                List<Entity> listCreateIns = new List<Entity>();
                //decimal phiBaoTriPaid = GetPhiBaoTri(enHD);
                decimal phiBaoTriPaid = 0;

                listCalendar = GetHoliday();

                //int bsd_typeunit = ((OptionSetValue)enHD["bsd_typeunit"]).Value;
                //traceS.Trace("bsd_typeunit " + bsd_typeunit);

                if (!enHD.Contains("bsd_paymentscheme"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_payment_scheme"));
                EntityReference refPS = (EntityReference)enHD["bsd_paymentscheme"];
                //if (bsd_typeunit == 100000000)//thấp tầng
                //{
                //    if (!enHD.Contains("bsd_paymentscheme"))
                //        throw new InvalidPluginExecutionException("Vui lòng chọn tiến độ thanh toán.");
                //    refPS = (EntityReference)enHD["bsd_paymentscheme"];
                //}
                //else
                //{
                //    if (!enHD.Contains("bsd_paymentschemeland"))
                //        throw new InvalidPluginExecutionException("Vui lòng chọn tiến độ thanh toán.");
                //    refPS = (EntityReference)enHD["bsd_paymentschemeland"];
                //}

                Entity enPS = service.Retrieve(refPS.LogicalName, refPS.Id, new ColumnSet(new string[] { "bsd_interestratemaster", "bsd_name" }));
                Entity interate = service.Retrieve(((EntityReference)enPS["bsd_interestratemaster"]).LogicalName, ((EntityReference)enPS["bsd_interestratemaster"]).Id,
                        new ColumnSet(new string[] { "bsd_gracedays", "bsd_depositinterest", "bsd_basecontractinterest" }));
                int graceday = interate.Contains("bsd_gracedays") ? (int)interate["bsd_gracedays"] : 0;
                decimal depositInterest = interate.Contains("bsd_depositinterest") ? (decimal)interate["bsd_depositinterest"] : 0;
                decimal baseContractInterest = interate.Contains("bsd_basecontractinterest") ? (decimal)interate["bsd_basecontractinterest"] : 0;
                //decimal interestPercent = interate.Contains("bsd_basecontractinterest") ? (decimal)interate["bsd_basecontractinterest"] : 0;

                //if (bsd_typeunit == 100000000)//thấp tầng
                //{
                //    int bsd_loaibangtinhgia = enHD.Contains("bsd_loaibangtinhgia") ? ((OptionSetValue)enHD["bsd_loaibangtinhgia"]).Value : -999;
                //    traceS.Trace("bsd_loaibangtinhgia " + bsd_loaibangtinhgia);
                //    switch (bsd_loaibangtinhgia)
                //    {
                //        case 100000006: //đất
                //            GenPaymentScheme(ref enHD, enPS, 100000000, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//đất
                //            break;
                //        case 100000007: //đất móng
                //            GenPaymentScheme(ref enHD, enPS, 100000000, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//đất
                //            GenPaymentScheme(ref enHD, enPS, 100000001, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//móng
                //            break;
                //        case 100000005: //đất nhà
                //            GenPaymentScheme(ref enHD, enPS, 100000000, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//đất
                //            GenPaymentScheme(ref enHD, enPS, 100000002, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//nhà
                //            break;
                //    }
                //}
                //else //cao tầng
                //{
                //    GenPaymentScheme(ref enHD, enPS, 100000003, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//cao tầng
                //}

                //GenPaymentScheme(ref enHD, enPS, 100000003, phiBaoTriPaid, ref listCreateIns, graceday, interestPercent);//cao tầng
                GenPaymentScheme(ref enHD, enPS, 100000002, phiBaoTriPaid, ref listCreateIns, graceday, depositInterest, baseContractInterest);//cao tầng


                if (listCreateIns.Count > 0)
                    BulkCreate(listCreateIns);
                traceS.Trace("5");

                #region task 1271
                listCreateIns = listCreateIns
                                .OrderBy(e => (e.Contains("bsd_pricetype") ? ((OptionSetValue)e["bsd_pricetype"]).Value : -1))
                                .ThenByDescending(e => (e.Contains("bsd_ordernumber") ? (int)e["bsd_ordernumber"] : -1))
                                .ToList();
                if (listCreateIns != null && listCreateIns.Count > 0)
                {
                    for (int i = 0; i < listCreateIns.Count - 1; i++)
                    {
                        if (((OptionSetValue)listCreateIns[i]["bsd_pricetype"]).Value == ((OptionSetValue)listCreateIns[i + 1]["bsd_pricetype"]).Value
                            && RetrieveLocalTimeFromUTCTime((DateTime)listCreateIns[i]["bsd_duedate"], service).Date < RetrieveLocalTimeFromUTCTime((DateTime)listCreateIns[i + 1]["bsd_duedate"], service).Date)
                        {
                            traceS.Trace(((OptionSetValue)listCreateIns[i + 1]["bsd_pricetype"]).Value + " || " + RetrieveLocalTimeFromUTCTime((DateTime)listCreateIns[i + 1]["bsd_duedate"], service));
                            throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "ins_invalid"));
                        }
                    }
                }
                #endregion
                traceS.Trace("6");

                decimal totalAmountPaid = 0;
                decimal depositFeePaid = 0;
                if (enHD.Contains("bsd_totalamountpaid"))
                {
                    totalAmountPaid = ((Money)enHD["bsd_totalamountpaid"]).Value;

                    if (enHD.Contains("bsd_depositfee"))
                    {
                        decimal depositFee = ((Money)enHD["bsd_depositfee"]).Value;
                        depositFeePaid = totalAmountPaid - depositFee < 0 ? totalAmountPaid : depositFee;
                    }
                }

                List<Entity> listUpdateIns = new List<Entity>();
                UpdateMoney(enHD, depositFeePaid, totalAmountPaid, ref listUpdateIns);

                if (listUpdateIns.Count > 0)
                    BulkUpdate(listUpdateIns);

                traceS.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        //private void GenPaymentScheme(ref Entity enHD, Entity enPS, int type, decimal phiBaoTriPaid, ref List<Entity> listCreateIns, int graceday, decimal interestPercent)
        private void GenPaymentScheme(ref Entity enHD, Entity enPS, int type, decimal phiBaoTriPaid, ref List<Entity> listCreateIns, int graceday, decimal depositInterest, decimal baseContractInterest)
        {
            traceS.Trace("vào GenPaymentScheme");
            decimal sumper = 0;
            decimal sumamount = 0;
            //bsd_freightamount = enHD.Contains("bsd_freightamount") ? ((Money)enHD["bsd_freightamount"]).Value : 0;
            //bsd_managementfee = enHD.Contains("bsd_managementfee") ? ((Money)enHD["bsd_managementfee"]).Value : 0;
            decimal amountCalcIns = enHD.Contains("bsd_totalamount") ? ((Money)enHD["bsd_totalamount"]).Value : 0;

            traceS.Trace("4.2");

            QueryExpression q = new QueryExpression("bsd_paymentschemedetailmaster");
            q.ColumnSet = new ColumnSet(true);
            q.AddOrder("bsd_ordernumber", OrderType.Ascending);
            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddCondition(new ConditionExpression("bsd_paymentscheme", ConditionOperator.Equal, enPS.Id));
            filter.AddCondition(new ConditionExpression("bsd_pricetype", ConditionOperator.Equal, type));
            filter.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            q.Criteria = filter;
            EntityCollection listInsMaster = service.RetrieveMultiple(q);

            int len = listInsMaster.Entities.Count;
            int orderNumber = 0;
            if (len == 0)
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_ins_psname", new Dictionary<string, object>
                {
                    ["ps_name"] = (string)enPS["bsd_name"]
                }));
            traceS.Trace("4.3");

            int cntInsValueNull = 0;
            int indexInsValueNull = 0;
            decimal sumValueNotNull = 0;
            decimal valuePer = 0;
            int typeGen = GetTypeGen(enPS, type, amountCalcIns, ref cntInsValueNull, ref sumValueNotNull, ref valuePer);

            bool f_lastinstallment = false;
            bool f_es = false;
            bool f_ESmaintenancefees = false;
            bool f_ESmanagementfee = false;
            bool isSPA = false;
            int i_dueCalMethod = -1;
            traceS.Trace("4.4");
            bool gotEstimateDate = false;
            DateTime? d_estimate = null;
            EntityCollection wordTemplateList = GetDinhNghiaWordTemplate(enPS);

            for (int i = 0; i < len; i++) // len = so luong INS detail
            {
                bool isLastIns = i == listInsMaster.Entities.Count - 1 ? true : false;
                f_ESmaintenancefees = listInsMaster.Entities[i].Contains("bsd_maintenancefees") ? (bool)listInsMaster.Entities[i]["bsd_maintenancefees"] : false;
                //f_ESmanagementfee = listInsMaster.Entities[i].Contains("bsd_managementfee") ? (bool)listInsMaster.Entities[i]["bsd_managementfee"] : false;

                if (listInsMaster.Entities[i].Contains("bsd_pinkbookhandover"))
                {
                    f_es = (bool)listInsMaster.Entities[i]["bsd_pinkbookhandover"];
                }
                else f_es = false;
                if (listInsMaster.Entities[i].Contains("bsd_lastinstallment") && (bool)listInsMaster.Entities[i]["bsd_lastinstallment"] == true)
                    f_lastinstallment = true;
                else f_lastinstallment = false;

                //traceS.Trace("4.5");

                if (!listInsMaster.Entities[i].Contains("bsd_duedatecalculatingmethod"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_duedate_calculating_method", new Dictionary<string, object>
                    {
                        ["ins_name"] = (string)listInsMaster.Entities[i]["bsd_name"],
                        ["ps_name"] = (string)enPS["bsd_name"]
                    }));
                i_dueCalMethod = listInsMaster.Entities[i].Contains("bsd_duedatecalculatingmethod") ? ((OptionSetValue)listInsMaster.Entities[i]["bsd_duedatecalculatingmethod"]).Value : 0;

                if (i_dueCalMethod == 100000002 && !gotEstimateDate) //Estimate handove date
                {
                    d_estimate = get_EstimatehandoverDate(enHD);
                    gotEstimateDate = true;
                    traceS.Trace($"d_estimate {d_estimate}");
                }

                if (i_dueCalMethod == 100000001) // auto
                {
                    //traceS.Trace("4.6");

                    //CreatePaymentPhase(enPS, ref orderNumber, listInsMaster.Entities[i], enHD, amountCalcIns,
                    //    f_ESmaintenancefees, f_ESmanagementfee, bsd_managementfee, bsd_freightamount, type, ref sumper,
                    //    ref sumamount, amountBeforeDiscount, isLastIns, phiBaoTriPaid, graceday, interestPercent, amountVATBeforeDiscount, amountVATAfterDiscount, ref sumVAT, typeGen,
                    //    ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer, ref listCreateIns, listInsMaster);
                    CreatePaymentPhase(enPS, ref orderNumber, listInsMaster.Entities[i], enHD, amountCalcIns, f_ESmaintenancefees, f_ESmanagementfee,
                        bsd_managementfee, bsd_freightamount, type, ref sumper, ref sumamount, isLastIns, phiBaoTriPaid, graceday, typeGen,
                        ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer, ref listCreateIns, listInsMaster, wordTemplateList,
                        ref isSPA, depositInterest, baseContractInterest);
                }
                else if (i_dueCalMethod == 100000000 || i_dueCalMethod == 100000002) // fixx
                {
                    //traceS.Trace("4.9.1");

                    //CreatePaymentPhase_fixDate(ref orderNumber, bsd_managementfee, bsd_freightamount, listInsMaster.Entities[i], enHD,
                    //    amountCalcIns, f_lastinstallment, f_es, f_ESmaintenancefees, f_ESmanagementfee, type, ref sumper, ref sumamount,
                    //    amountBeforeDiscount, isLastIns, phiBaoTriPaid, graceday, interestPercent, amountVATBeforeDiscount, amountVATAfterDiscount, ref sumVAT, typeGen,
                    //    ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer, ref listCreateIns);
                    CreatePaymentPhase_fixDate(ref orderNumber, bsd_managementfee, bsd_freightamount, listInsMaster.Entities[i], enHD,
                        amountCalcIns, f_lastinstallment, f_es, f_ESmaintenancefees, f_ESmanagementfee, type, ref sumper, ref sumamount,
                        isLastIns, phiBaoTriPaid, graceday, typeGen, ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer, ref listCreateIns,
                        i_dueCalMethod, d_estimate, wordTemplateList, ref isSPA, depositInterest, baseContractInterest);
                }
            }
            traceS.Trace("xong GenPaymentScheme");
        }

        private decimal GetTax(EntityReference taxcode)
        {
            Entity tax = service.Retrieve(taxcode.LogicalName, taxcode.Id, new ColumnSet(new string[] { "bsd_name", "bsd_value" }));
            if (!tax.Attributes.Contains("bsd_value"))
                throw new InvalidPluginExecutionException("Please input tax value!");
            return (decimal)tax["bsd_value"];
        }

        //private void CreatePaymentPhase(Entity PM, ref int orderNumber, Entity en, Entity enHD, decimal amountCalcIns,
        //    bool f_ESmaintenancefees, bool f_ESmanagementfee, decimal bsd_managementfee, decimal bsd_maintenancefees, int typePrice,
        //    ref decimal sumper, ref decimal sumamount, decimal amountBeforeDiscount, bool isLastIns, decimal phiBaoTriPaid, int graceday,
        //    decimal interestPercent, decimal amountVATBeforeDiscount, decimal amountVATAfterDiscount, ref decimal sumVAT, int typeGen, ref int cntInsValueNull, ref decimal sumValueNotNull,
        //    ref int indexInsValueNull, ref decimal valuePer, ref List<Entity> listCreateIns, EntityCollection listInsMaster)
        private void CreatePaymentPhase(Entity enPS, ref int orderNumber, Entity en, Entity enHD, decimal amountCalcIns, bool f_ESmaintenancefees,
        bool f_ESmanagementfee, decimal bsd_managementfee, decimal bsd_maintenancefees, int typePrice, ref decimal sumper, ref decimal sumamount, bool isLastIns,
        decimal phiBaoTriPaid, int graceday, int typeGen, ref int cntInsValueNull, ref decimal sumValueNotNull, ref int indexInsValueNull, ref decimal valuePer,
        ref List<Entity> listCreateIns, EntityCollection listInsMaster, EntityCollection wordTemplateList, ref bool isSPA, decimal depositInterest, decimal baseContractInterest)
        {
            traceS.Trace("vào CreatePaymentPhase");
            orderNumber++;

            if (en.Contains("bsd_startingwith") && orderNumber != 1)
            {
                if (!en.Contains("bsd_nextperiodtype"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_nextperiodtype", new Dictionary<string, object>
                    {
                        ["ins_name"] = en["bsd_name"].ToString(),
                        ["ps_name"] = (string)enPS["bsd_name"]
                    }));

                int type = ((OptionSetValue)en["bsd_nextperiodtype"]).Value;
                if (type == 1)//month
                {
                    if (!en.Attributes.Contains("bsd_numberofnextmonth"))
                        throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_numberofnextmonth", new Dictionary<string, object>
                        {
                            ["ins_name"] = en["bsd_name"].ToString(),
                            ["ps_name"] = (string)enPS["bsd_name"]
                        }));
                }
                else if (type == 2)//day
                {
                    if (!en.Attributes.Contains("bsd_numberofnextdays"))
                        throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_numberofnextdays", new Dictionary<string, object>
                        {
                            ["ins_name"] = en["bsd_name"].ToString(),
                            ["ps_name"] = (string)enPS["bsd_name"]
                        }));
                }
            }

            //traceS.Trace("orderNumber " + orderNumber);
            Entity tmp = new Entity("bsd_paymentschemedetail");
            tmp["bsd_ordernumber"] = orderNumber;
            tmp["bsd_name"] = "Đợt " + orderNumber;
            //tmp["bsd_code"] = string.Format("{0}-{1:ddMMyyyyhhmmssff}", tmp["bsd_name"], DateTime.UtcNow);
            tmp["bsd_reservationcontract"] = enHD.ToEntityReference();
            tmp["bsd_paymentscheme"] = en["bsd_paymentscheme"];
            //tmp["bsd_project"] = en.Contains("bsd_project") ? en["bsd_project"] : null;
            tmp["bsd_project"] = enHD.Contains("bsd_projectid") ? enHD["bsd_projectid"] : null;
            tmp["bsd_amountwaspaid"] = new Money(0);
            tmp["bsd_pricetype"] = new OptionSetValue(typePrice);
            tmp["bsd_gracedays"] = graceday;
            //tmp["bsd_interestpercent"] = interestPercent;
            tmp["bsd_calculationmethodmaster"] = en.Contains("bsd_calculationmethod") ? en["bsd_calculationmethod"] : null;
            tmp["bsd_amount"] = en.Contains("bsd_amount") ? en["bsd_amount"] : null;

            bool signContractInstallment = en.Contains("bsd_signcontractinstallment") ? (bool)en["bsd_signcontractinstallment"] : false;
            tmp["bsd_signcontractinstallment"] = signContractInstallment;
            tmp["bsd_pinkbookhandover"] = en.Contains("bsd_pinkbookhandover") ? en["bsd_pinkbookhandover"] : false;
            tmp["bsd_lastinstallment"] = en.Contains("bsd_lastinstallment") ? en["bsd_lastinstallment"] : false;
            tmp["bsd_official"] = en.Contains("bsd_official") ? en["bsd_official"] : false;
            tmp["bsd_gopdot"] = en.Contains("bsd_gopdot") ? en["bsd_gopdot"] : false;

            if (!isSPA && signContractInstallment)
                isSPA = true;
            if (isSPA)  //SPA
                tmp["bsd_interestpercent"] = baseContractInterest;
            else  //EDA
                tmp["bsd_interestpercent"] = depositInterest;

            tmp["bsd_duedate"] = calculateDuedate(enHD, en, listCreateIns, listInsMaster, orderNumber == 1);

            tmp["bsd_calendartype"] = en.Contains("bsd_calendartype") ? en["bsd_calendartype"] : null;
            tmp["bsd_waiverinterest"] = new Money(0);

            tmp["bsd_actualgracedays"] = 0;

            tmp["bsd_interestwaspaid"] = new Money(0);
            tmp["bsd_interestchargeamount"] = new Money(0);

            tmp["bsd_maintenancefeepaid"] = new Money(0);
            tmp["bsd_maintenanceamount"] = new Money(0);
            tmp["bsd_maintenancefeewaiver"] = new Money(0);

            tmp["bsd_depositamount"] = new Money(0);

            #region  extra field
            tmp["bsd_startfrominstallment"] = en.Contains("bsd_startfrominstallment") ? en["bsd_startfrominstallment"] : null; // 29/6/2023

            if (en.Contains("bsd_nextperiodtype"))
            {
                int bsd_nextperiodtype = ((OptionSetValue)en["bsd_nextperiodtype"]).Value;
                tmp["bsd_nextperiodtype"] = new OptionSetValue(bsd_nextperiodtype);
            }

            if (en.Contains("bsd_numberofnextdays"))
            {
                double bsd_numberofnextdays = double.Parse(en["bsd_numberofnextdays"].ToString());
                tmp["bsd_numberofnextdays"] = bsd_numberofnextdays;
            }

            if (en.Contains("bsd_numberofnextmonth"))
            {
                int bsd_numberofnextmonth = (int)en["bsd_numberofnextmonth"];
                tmp["bsd_numberofnextmonth"] = bsd_numberofnextmonth;
            }

            if (en.Contains("bsd_typepayment"))
            {
                int i_bsd_typepayment = ((OptionSetValue)en["bsd_typepayment"]).Value;
                tmp["bsd_typepayment"] = new OptionSetValue(i_bsd_typepayment);
            }

            if (en.Contains("bsd_number"))
            {
                int bsd_number = (int)en["bsd_number"];
                tmp["bsd_number"] = bsd_number;
            }
            #endregion

            decimal tmpamount = 0;
            CalcAmount(ref tmpamount, ref tmp, en, typeGen, isLastIns, amountCalcIns,
            ref sumper, ref sumamount, ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer);

            tmp["bsd_amountofthisphase"] = new Money(tmpamount);
            tmp["bsd_balance"] = new Money(tmpamount);
            tmp["bsd_duedatecalculatingmethod"] = new OptionSetValue(100000001);


            #region if bsd_maintenancefees/ bsd_managementfee = yes => set amount
            tmp["bsd_maintenancefees"] = f_ESmaintenancefees;
            //tmp["bsd_managementfee"] = f_ESmanagementfee;

            //tmp["bsd_managementamount"] = f_ESmanagementfee ? new Money(bsd_managementfee) : new Money(0);
            tmp["bsd_maintenanceamount"] = f_ESmaintenancefees ? new Money(bsd_maintenancefees) : new Money(0);
            tmp["bsd_maintenancefeepaid"] = f_ESmaintenancefees ? new Money(phiBaoTriPaid) : new Money(0);
            #endregion
            tmp["bsd_originalduedate"] = tmp["bsd_duedate"];
            SetTextWordTemplate(ref tmp, wordTemplateList, orderNumber);
            tmp.Id = Guid.NewGuid();
            //service.Create(tmp);
            listCreateIns.Add(tmp);

            //traceS.Trace("ra CreatePaymentPhase");
        }

        //private void CreatePaymentPhase_fixDate(ref int orderNumber, decimal bsd_managementfee, decimal bsd_maintenancefees, Entity en, Entity enHD,
        //    decimal amountCalcIns, bool f_last, bool f_es, bool f_ESmaintenancefees, bool f_ESmanagementfee, int typePrice, ref decimal sumper,
        //    ref decimal sumamount, decimal amountBeforeDiscount, bool isLastIns, decimal phiBaoTriPaid, int graceday, decimal interestPercent,
        //    decimal amountVATBeforeDiscount, decimal amountVATAfterDiscount, ref decimal sumVAT, int typeGen, ref int cntInsValueNull, ref decimal sumValueNotNull, ref int indexInsValueNull,
        //    ref decimal valuePer, ref List<Entity> listCreateIns)
        private void CreatePaymentPhase_fixDate(ref int orderNumber, decimal bsd_managementfee, decimal bsd_maintenancefees, Entity en, Entity enHD,
        decimal amountCalcIns, bool f_last, bool f_es, bool f_ESmaintenancefees, bool f_ESmanagementfee, int typePrice, ref decimal sumper,
        ref decimal sumamount, bool isLastIns, decimal phiBaoTriPaid, int graceday, int typeGen, ref int cntInsValueNull, ref decimal sumValueNotNull,
        ref int indexInsValueNull, ref decimal valuePer, ref List<Entity> listCreateIns, int i_dueCalMethod, DateTime? d_estimate, EntityCollection wordTemplateList,
        ref bool isSPA, decimal depositInterest, decimal baseContractInterest)
        {
            traceS.Trace("vào CreatePaymentPhase_fixDate");
            Entity tmp = new Entity("bsd_paymentschemedetail");
            orderNumber++;
            //traceS.Trace(en["bsd_name"] + " " + en.Id);

            tmp["bsd_ordernumber"] = orderNumber;
            tmp["bsd_name"] = "Đợt " + orderNumber;
            //tmp["bsd_code"] = string.Format("{0}-{1:ddMMyyyyhhmmssff}", tmp["bsd_name"], DateTime.UtcNow);
            tmp["bsd_reservationcontract"] = enHD.ToEntityReference();
            tmp["bsd_paymentscheme"] = en["bsd_paymentscheme"];
            //tmp["bsd_project"] = en.Contains("bsd_project") ? en["bsd_project"] : null;
            tmp["bsd_project"] = enHD.Contains("bsd_projectid") ? enHD["bsd_projectid"] : null;
            tmp["bsd_amountwaspaid"] = new Money(0);
            tmp["bsd_depositamount"] = new Money(0);
            tmp["bsd_pricetype"] = new OptionSetValue(typePrice);
            tmp["bsd_waiverinterest"] = new Money(0);
            tmp["bsd_gracedays"] = graceday;
            //tmp["bsd_interestpercent"] = interestPercent;
            tmp["bsd_calculationmethodmaster"] = en.Contains("bsd_calculationmethod") ? en["bsd_calculationmethod"] : null;
            tmp["bsd_amount"] = en.Contains("bsd_amount") ? en["bsd_amount"] : null;

            tmp["bsd_actualgracedays"] = 0;
            tmp["bsd_calendartype"] = en.Contains("bsd_calendartype") ? en["bsd_calendartype"] : null;
            tmp["bsd_interestwaspaid"] = new Money(0);
            tmp["bsd_interestchargeamount"] = new Money(0);

            tmp["bsd_maintenancefeepaid"] = new Money(0);
            tmp["bsd_maintenanceamount"] = new Money(0);
            tmp["bsd_maintenancefeewaiver"] = new Money(0);

            bool signContractInstallment = en.Contains("bsd_signcontractinstallment") ? (bool)en["bsd_signcontractinstallment"] : false;
            tmp["bsd_signcontractinstallment"] = signContractInstallment;
            tmp["bsd_pinkbookhandover"] = en.Contains("bsd_pinkbookhandover") ? en["bsd_pinkbookhandover"] : false;
            tmp["bsd_lastinstallment"] = en.Contains("bsd_lastinstallment") ? en["bsd_lastinstallment"] : false;
            tmp["bsd_official"] = en.Contains("bsd_official") ? en["bsd_official"] : false;
            tmp["bsd_gopdot"] = en.Contains("bsd_gopdot") ? en["bsd_gopdot"] : false;

            if (!isSPA && signContractInstallment)
                isSPA = true;
            if (isSPA)  //SPA
                tmp["bsd_interestpercent"] = baseContractInterest;
            else  //EDA
                tmp["bsd_interestpercent"] = depositInterest;

            decimal tmpamount = 0;
            CalcAmount(ref tmpamount, ref tmp, en, typeGen, isLastIns, amountCalcIns,
            ref sumper, ref sumamount, ref cntInsValueNull, ref sumValueNotNull, ref indexInsValueNull, ref valuePer);

            tmp["bsd_duedatecalculatingmethod"] = new OptionSetValue(i_dueCalMethod);
            if (i_dueCalMethod == 100000002)    //Estimate handove date
            {
                tmp["bsd_duedate"] = d_estimate;
                tmp["bsd_fixeddate"] = d_estimate;
            }
            else
            {
                if (en.Contains("bsd_fixeddate"))
                {
                    tmp["bsd_duedate"] = en["bsd_fixeddate"];
                    tmp["bsd_fixeddate"] = en["bsd_fixeddate"];
                }
            }

            tmp["bsd_amountofthisphase"] = new Money(tmpamount);
            tmp["bsd_balance"] = new Money(tmpamount);

            #region if bsd_maintenancefees/ bsd_managementfee = yes => set amount
            tmp["bsd_maintenancefees"] = f_ESmaintenancefees;
            //tmp["bsd_managementfee"] = f_ESmanagementfee;

            //tmp["bsd_managementamount"] = f_ESmanagementfee ? new Money(bsd_managementfee) : new Money(0);
            tmp["bsd_maintenanceamount"] = f_ESmaintenancefees ? new Money(bsd_maintenancefees) : new Money(0);
            tmp["bsd_maintenancefeepaid"] = f_ESmaintenancefees ? new Money(phiBaoTriPaid) : new Money(0);
            #endregion
            tmp["bsd_originalduedate"] = tmp["bsd_duedate"];
            SetTextWordTemplate(ref tmp, wordTemplateList, orderNumber);
            tmp.Id = Guid.NewGuid();

            //service.Create(tmp);
            listCreateIns.Add(tmp);

            //traceS.Trace("ra CreatePaymentPhase_fixDate");
        }

        private DateTime calculateDuedate(Entity enHD, Entity enIns, List<Entity> listCreateIns, EntityCollection listInsMaster, bool isFirstIns)
        {
            traceS.Trace("calculateDuedate đầu function");
            bool flag = false;
            DateTime date = DateTime.UtcNow;

            if (isFirstIns)
            {
                if (!enHD.Contains("bsd_racontractsigndate"))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_racontractsigndate"));
                date = CalcDate(enIns, (DateTime)enHD["bsd_racontractsigndate"], ref flag);
            }
            else
            {
                Entity fromIns = listInsMaster.Entities.FirstOrDefault(e => e.Id == ((EntityReference)enIns["bsd_startfrominstallment"]).Id);
                int bsd_pricetype = ((OptionSetValue)fromIns["bsd_pricetype"]).Value;
                int bsd_ordernumber = (int)fromIns["bsd_ordernumber"];

                Entity item = listCreateIns.FirstOrDefault(e => e.Contains("bsd_pricetype") && ((OptionSetValue)e["bsd_pricetype"]).Value == bsd_pricetype &&
                                    e.Contains("bsd_ordernumber") && (int)e["bsd_ordernumber"] == bsd_ordernumber);

                date = CalcDate(enIns, (DateTime)item["bsd_duedate"], ref flag);
            }


            if (!flag)
            {
                if (enIns.Contains("bsd_nextperiodtype"))
                {
                    //Month   1
                    //Day 2
                    int bsd_nextperiodtype = ((OptionSetValue)enIns["bsd_nextperiodtype"]).Value;
                    if (bsd_nextperiodtype == 1)
                    {
                        date = date.AddMonths((int)enIns["bsd_numberofnextmonth"]);
                    }
                    else if (bsd_nextperiodtype == 2)
                    {
                        date = CalculateWorkday(date, double.Parse(enIns["bsd_numberofnextdays"].ToString()), enIns);
                    }
                }
                else
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "ins_invalid"));
            }
            //traceS.Trace("calculateDuedate cuối function");
            return date;
        }

        private DateTime CalcDate(Entity enIns, DateTime dateSign, ref bool flag)
        {
            traceS.Trace("CalcDate");
            DateTime date = DateTime.Now;
            //Month   1
            //Day 2
            int bsd_nextperiodtype = ((OptionSetValue)enIns["bsd_nextperiodtype"]).Value;
            if (bsd_nextperiodtype == 1)
            {
                date = RetrieveLocalTimeFromUTCTime(dateSign, service).AddMonths((int)enIns["bsd_numberofnextmonth"]);
                flag = true;
            }
            else if (bsd_nextperiodtype == 2)
            {
                date = CalculateWorkday(RetrieveLocalTimeFromUTCTime(dateSign, service), double.Parse(enIns["bsd_numberofnextdays"].ToString()), enIns);
                flag = true;
            }

            return date;
        }

        private DateTime get_EstimatehandoverDate(Entity enHD)
        {
            if (!enHD.Contains("bsd_unitno"))
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_unitinformation"));

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_product"">
                <attribute name=""bsd_estimatehandoverdate"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_productid"" operator=""eq"" value=""{((EntityReference)enHD["bsd_unitno"]).Id}"" />
                  <condition attribute=""bsd_estimatehandoverdate"" operator=""not-null"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                return (DateTime)rs.Entities[0]["bsd_estimatehandoverdate"];
            else
                return get_EstimateFromProject(enHD);
        }

        private DateTime get_EstimateFromProject(Entity enHD)
        {
            if (!enHD.Contains("bsd_projectid"))
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_project"));

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch top=""1"">
              <entity name=""bsd_project"">
                <attribute name=""bsd_estimatehandoverdate"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_projectid"" operator=""eq"" value=""{((EntityReference)enHD["bsd_projectid"]).Id}"" />
                  <condition attribute=""bsd_estimatehandoverdate"" operator=""not-null"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
                return (DateTime)rs.Entities[0]["bsd_estimatehandoverdate"];
            else
                throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "no_estimatehandoverdate"));
        }

        public DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, IOrganizationService ser)
        {
            int? timeZoneCode = RetrieveCurrentUsersSettings(ser);
            if (!timeZoneCode.HasValue)
                throw new InvalidPluginExecutionException("Can't find time zone code");
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = utcTime.ToUniversalTime()
            };

            LocalTimeFromUtcTimeResponse response = (LocalTimeFromUtcTimeResponse)ser.Execute(request);
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

        private List<DateTime> GetHoliday()
        {
            traceS.Trace("GetHolidays");
            List<DateTime> holidays = new List<DateTime>();
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""calendar"">
                <filter>
                  <condition attribute=""type"" operator=""eq"" value=""2"" />
                </filter>
                <link-entity name=""calendarrule"" from=""calendarid"" to=""calendarid"">
                  <attribute name=""name"" alias=""namerule"" />
                  <attribute name=""effectiveintervalstart"" alias=""startrule"" />
                  <attribute name=""effectiveintervalend"" alias=""endrule"" />
                  <attribute name=""duration"" alias=""durationrule"" />
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities.Count > 0)
            {
                foreach (var item in rs.Entities)
                {
                    int duration = item.Contains("durationrule") ? ((int)((AliasedValue)item["durationrule"]).Value / 1440) : 0;
                    DateTime dateHolidayStart = RetrieveLocalTimeFromUTCTime((DateTime)((AliasedValue)item["startrule"]).Value, service);
                    for (int i = 0; i < duration; i++)
                    {
                        holidays.Add(dateHolidayStart.AddDays(i).Date);
                    }
                }
            }

            holidays = holidays.Distinct().ToList();
            return holidays;
        }

        public DateTime CalculateWorkday(DateTime startDate, double days, Entity ins)
        {
            //traceS.Trace("CalculateWorkday");
            if (days == 0)
                return startDate;
            int bsd_calendartype = ins.Contains("bsd_calendartype") ? ((OptionSetValue)ins["bsd_calendartype"]).Value : 0;
            //traceS.Trace("startDate " + startDate + " " + bsd_calendartype);
            DateTime resultDate = startDate;

            DateTime? convertedDate = CalculateWorkday_New(startDate, (int)days, bsd_calendartype);
            if (convertedDate != null)
            {
                resultDate = (DateTime)convertedDate;
            }
            return resultDate;
        }

        private DateTime? CalculateWorkday_New(DateTime startDate, int workDaysToAdd, int workingDayType)
        {
            DateTime? convertedDate = null;
            OrganizationRequest req = new OrganizationRequest("bsd_Action_CalculateWorkday");
            req["startDate"] = startDate.ToString("d/M/yyyy");
            req["workDaysToAdd"] = workDaysToAdd;
            req["workingDayType"] = workingDayType;
            if (listCalendar.Count > 0)
                req["holidays"] = string.Join(",", listCalendar);
            OrganizationResponse response = service.Execute(req);
            if (response.Results.Count > 0 && response.Results.Contains("resultDate") && response.Results["resultDate"] != null)
            {
                convertedDate = (DateTime)response.Results["resultDate"];
            }
            //traceS.Trace("convertedDate " + convertedDate);
            return convertedDate;
        }

        private void UpdateMoney(Entity enHD, decimal depositFeePaid, decimal totalAmountPaid, ref List<Entity> listUpdateIns)
        {
            traceS.Trace("UpdateMoney");
            var fetchXmlKhongGop = $@"
                        <fetch>
                          <entity name='bsd_paymentschemedetail'>
                            <attribute name='bsd_name' />
                            <attribute name='bsd_ordernumber' />
                            <attribute name='bsd_pricetype' />
                            <attribute name='bsd_amountofthisphase' />
                            <attribute name='bsd_maintenancefeepaid' />
                            <attribute name='bsd_maintenancefeeremaining' />
                            <filter type='and'>
                              <condition attribute='bsd_reservationcontract' operator='eq' value='{enHD.Id}'/>
                              <condition attribute='statecode' operator='eq' value='0'/>
                            </filter>
                            <order attribute='bsd_duedate' />
                            <order attribute='bsd_ordernumber' />
                            <order attribute=""bsd_pricetype"" />
                          </entity>
                        </fetch>";
            var lstKhongGop = service.RetrieveMultiple(new FetchExpression(fetchXmlKhongGop.ToString()));
            decimal amountwaspaid = totalAmountPaid;
            foreach (var item in lstKhongGop.Entities)
            {
                UpdateMoney_Detail(item, depositFeePaid, ref amountwaspaid, ref listUpdateIns);
            }
        }

        private void UpdateMoney_Detail(Entity item, decimal depositFeePaid, ref decimal amountwaspaid, ref List<Entity> listUpdateIns)
        {

            Entity enUpdate = new Entity(item.LogicalName, item.Id);
            #region reset
            enUpdate["bsd_depositamount"] = new Money(0);
            enUpdate["bsd_amountwaspaid"] = new Money(0);
            enUpdate["statuscode"] = new OptionSetValue(100000000);
            enUpdate["bsd_balance"] = item["bsd_amountofthisphase"];
            #endregion

            decimal bsd_amountofthisphase = item.Contains("bsd_amountofthisphase") ? ((Money)item["bsd_amountofthisphase"]).Value : 0;
            decimal tmp = bsd_amountofthisphase - amountwaspaid;
            if ((int)item["bsd_ordernumber"] == 1)
            {
                enUpdate["bsd_depositamount"] = new Money(depositFeePaid);
            }

            if (tmp < 0)
            {
                enUpdate["bsd_amountwaspaid"] = new Money(bsd_amountofthisphase);
                enUpdate["statuscode"] = new OptionSetValue(100000001);
                enUpdate["bsd_balance"] = new Money(0);

                amountwaspaid = amountwaspaid - bsd_amountofthisphase;
            }
            else
            {
                if (tmp == 0)
                {
                    enUpdate["statuscode"] = new OptionSetValue(100000001);
                }

                enUpdate["bsd_amountwaspaid"] = new Money(amountwaspaid);
                enUpdate["bsd_balance"] = new Money(tmp);

                amountwaspaid = 0;
            }

            decimal bsd_maintenancefeepaid = item.Contains("bsd_maintenancefeepaid") ? ((Money)item["bsd_maintenancefeepaid"]).Value : 0;
            decimal bsd_maintenancefeeremaining = item.Contains("bsd_maintenancefeeremaining") ? ((Money)item["bsd_maintenancefeeremaining"]).Value : 0;
            enUpdate["bsd_maintenancefeesstatus"] = bsd_maintenancefeepaid > 0 && bsd_maintenancefeeremaining <= 0 ? true : false;
            //service.Update(enHandover);
            listUpdateIns.Add(enUpdate);
        }

        private decimal GetPhiBaoTri(Entity enHD)
        {
            traceS.Trace("GetPhiBaoTri");

            decimal sumPhiBaoTri = 0;
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch aggregate=""true"">
              <entity name=""bsd_payment"">
                <attribute name=""bsd_amountpay"" alias=""sumAmount"" aggregate=""sum"" />
                <filter>
                  <condition attribute=""bsd_paymenttype"" operator=""eq"" value=""100000002"" />
                  <condition attribute=""bsd_reservationcontract"" operator=""eq"" value=""{enHD.Id}"" />
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000000"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                if (((AliasedValue)rs[0]["sumAmount"]).Value != null)
                    sumPhiBaoTri = ((Money)((AliasedValue)rs[0]["sumAmount"]).Value).Value;
            }

            return sumPhiBaoTri;
        }

        private int GetTypeGen(Entity paymentScheme, int bsd_pricetype, decimal amountCalcIns, ref int cntInsValueNull,
            ref decimal sumValueNotNull, ref decimal valuePer)
        {
            traceS.Trace("GetTypeGen");

            int result = 0;     // 0: Default (full %), 1: full tiền, 2: % + null, 3: tiền + null

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_paymentschemedetailmaster"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_amount"" />
                <attribute name=""bsd_amountpercent"" />
                <attribute name=""bsd_ordernumber"" />
                <attribute name=""bsd_calculationmethod"" />
                <filter>
                  <condition attribute=""bsd_paymentscheme"" operator=""eq"" value=""{paymentScheme.Id}"" />
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_pricetype"" operator=""eq"" value=""{bsd_pricetype}"" />
                </filter>
                <order attribute=""bsd_ordernumber"" />
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                bool fullPercent = true;
                bool fullAmount = true;
                bool isTypeAmount = false;
                bool isTypePercent = false;

                foreach (var item in rs.Entities)
                {
                    bool hasAmount = item.Contains("bsd_amount") && item["bsd_amount"] != null;
                    bool hasPercent = item.Contains("bsd_amountpercent") && item["bsd_amountpercent"] != null;
                    int calculationMethod = item.Contains("bsd_calculationmethod") ? ((OptionSetValue)item["bsd_calculationmethod"]).Value : 0;
                    // chỉ có percent

                    if (calculationMethod == 100000000)
                        isTypePercent = true;

                    if (calculationMethod == 100000001)
                        isTypeAmount = true;

                    // Nếu cả 2 loại
                    if (isTypePercent && isTypeAmount)
                        throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "ins_include_amount_percent"));

                    if (isTypePercent && hasPercent)
                        sumValueNotNull += (decimal)item["bsd_amountpercent"];
                    else
                        fullPercent = false;

                    // chỉ có amount
                    if (isTypeAmount && hasAmount)
                        sumValueNotNull += ((Money)item["bsd_amount"]).Value;
                    else
                        fullAmount = false;

                    if ((isTypeAmount && !hasAmount) || (isTypePercent && !hasPercent))
                        cntInsValueNull += 1;

                }

                decimal amountDiscountTmp = amountCalcIns;

                if ((fullAmount && amountDiscountTmp < sumValueNotNull) || (isTypeAmount && amountDiscountTmp < sumValueNotNull))
                    throw new InvalidPluginExecutionException(MessageProvider.GetMessage(service, context, "ins_exceeding_total_value"));

                if (isTypeAmount && cntInsValueNull > 0)
                    valuePer = Math.Round((amountDiscountTmp - sumValueNotNull) / cntInsValueNull, MidpointRounding.AwayFromZero);
                else if (isTypePercent && cntInsValueNull > 0)
                    valuePer = Math.Round((100 - sumValueNotNull) / cntInsValueNull, 2, MidpointRounding.AwayFromZero);

                result = fullPercent ? 0 : fullAmount ? 1 : isTypePercent ? 2 : isTypeAmount ? 3 : 0;
            }

            traceS.Trace($"TypeGen {bsd_pricetype} || {result}");
            return result;
        }

        private decimal CalcAmountIns(bool isLastIns, decimal amountDiscountTmp, decimal sumamount, decimal sumperIncludeSelf)
        {
            decimal tmpAmount = 0;

            if (isLastIns)
            {
                tmpAmount = Math.Round(amountDiscountTmp - sumamount, MidpointRounding.AwayFromZero);
            }
            else
            {
                tmpAmount = Math.Round(((sumperIncludeSelf * amountDiscountTmp / 100) - sumamount), MidpointRounding.AwayFromZero);
            }

            if (tmpAmount < 0) tmpAmount = 0;
            return tmpAmount;
        }

        private decimal CalcPercentFromAmount(bool isLastIns, decimal amount, decimal amountDiscountTmp, decimal sumper)
        {
            decimal percent = 0;

            if (isLastIns)
            {
                percent = Math.Round(100 - sumper, 2, MidpointRounding.AwayFromZero);
            }
            else
            {
                percent = amountDiscountTmp > 0 ? Math.Round((amount / amountDiscountTmp) * 100, 2, MidpointRounding.AwayFromZero) : 0;
            }

            return percent;
        }

        private void CalcAmount(ref decimal tmpamount, ref Entity tmp, Entity ins, int typeGen, bool isLastIns, decimal amountCalcIns,
            ref decimal sumper, ref decimal sumamount, ref int cntInsValueNull, ref decimal sumValueNotNull, ref int indexInsValueNull, ref decimal valuePer)
        {
            decimal amountDiscountTmp = amountCalcIns;

            decimal percent = 0;
            switch (typeGen)
            {
                case 1: // full tiền
                    tmpamount = ((Money)ins["bsd_amount"]).Value;

                    percent = CalcPercentFromAmount(isLastIns, tmpamount, amountDiscountTmp, sumper);
                    sumper += percent;
                    tmp["bsd_amountpercent"] = percent;
                    break;
                case 2: //% + null
                    if (ins.Contains("bsd_amountpercent"))
                    {
                        percent = (decimal)ins["bsd_amountpercent"];
                    }
                    else
                    {
                        indexInsValueNull++;
                        if (cntInsValueNull == indexInsValueNull)   // đợt cuối của những đợt bsd_amountpercent = null
                            percent = Math.Round(100 - sumValueNotNull - (valuePer * (cntInsValueNull - 1)), 2, MidpointRounding.AwayFromZero);
                        else
                            percent = valuePer;
                    }

                    sumper += percent;
                    tmpamount = CalcAmountIns(isLastIns, amountDiscountTmp, sumamount, sumper);
                    tmp["bsd_amountpercent"] = percent;
                    break;
                case 3: //tiền + null
                    if (ins.Contains("bsd_amount"))
                    {
                        tmpamount = ((Money)ins["bsd_amount"]).Value;
                    }
                    else
                    {
                        indexInsValueNull++;
                        if (cntInsValueNull == indexInsValueNull)   // đợt cuối của những đợt bsd_amountpercent = null
                            tmpamount = Math.Round(amountDiscountTmp - sumValueNotNull - (valuePer * (cntInsValueNull - 1)), MidpointRounding.AwayFromZero);
                        else
                            tmpamount = valuePer;
                    }

                    percent = CalcPercentFromAmount(isLastIns, tmpamount, amountDiscountTmp, sumper);
                    sumper += percent;
                    tmp["bsd_amountpercent"] = percent;
                    break;
                default:    //full %
                    percent = (decimal)ins["bsd_amountpercent"];

                    sumper += percent;
                    tmpamount = CalcAmountIns(isLastIns, amountDiscountTmp, sumamount, sumper);
                    tmp["bsd_amountpercent"] = percent;
                    break;
            }

            sumamount += tmpamount;
            traceS.Trace($"info: {ins.Id} || amountCalcIns: {amountCalcIns} || sumamount: {sumamount} || sumper: {sumper} || tmpamount: {tmpamount} || valuePer: {valuePer} || {cntInsValueNull} || {indexInsValueNull}");
        }

        private void BulkCreate(List<Entity> entities)
        {
            traceS.Trace("BulkCreate");

            try
            {
                traceS.Trace($"requests {entities.Count}");

                ExecuteTransactionRequest transactionRequest = new ExecuteTransactionRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    ReturnResponses = false
                };

                foreach (var entity in entities)
                {
                    transactionRequest.Requests.Add(new CreateRequest { Target = entity });
                }

                service.Execute(transactionRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Transaction failed: {ex.Message}", ex);
            }
        }

        private void BulkUpdate(List<Entity> entities)
        {
            traceS.Trace("BulkUpdate");

            try
            {
                traceS.Trace($"requests {entities.Count}");

                ExecuteTransactionRequest transactionRequest = new ExecuteTransactionRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    ReturnResponses = false
                };

                foreach (var entity in entities)
                {
                    transactionRequest.Requests.Add(new UpdateRequest { Target = entity });
                }

                service.Execute(transactionRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Transaction failed: {ex.Message}", ex);
            }
        }

        private EntityCollection GetDinhNghiaWordTemplate(Entity paymentScheme)
        {
            traceS.Trace("GetDinhNghiaWordTemplate");

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
              <entity name=""bsd_dinhnghiawordtemplate"">
                <attribute name=""bsd_text1"" />
                <attribute name=""bsd_text2"" />
                <attribute name=""bsd_text3"" />
                <attribute name=""bsd_text4"" />
                <attribute name=""bsd_text5"" />
                <attribute name=""bsd_text6"" />
                <attribute name=""bsd_text7"" />
                <attribute name=""bsd_text8"" />
                <attribute name=""bsd_text9"" />
                <attribute name=""bsd_text10"" />
                <filter>
                  <condition attribute=""bsd_paymentscheme"" operator=""eq"" value=""{paymentScheme.Id}"" />
                </filter>
                <order attribute=""createdon"" />
              </entity>
            </fetch>";
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        private void SetTextWordTemplate(ref Entity tmp, EntityCollection wordTemplateList, int orderNumber)
        {
            traceS.Trace("SetTextWordTemplate");
            if (wordTemplateList != null && wordTemplateList.Entities.Count >= orderNumber)
            {
                Entity item = wordTemplateList[orderNumber - 1];
                tmp["bsd_text1"] = item.Contains("bsd_text1") ? item["bsd_text1"] : null;
                tmp["bsd_text2"] = item.Contains("bsd_text2") ? item["bsd_text2"] : null;
                tmp["bsd_text3"] = item.Contains("bsd_text3") ? item["bsd_text3"] : null;
                tmp["bsd_text4"] = item.Contains("bsd_text4") ? item["bsd_text4"] : null;
                tmp["bsd_text5"] = item.Contains("bsd_text5") ? item["bsd_text5"] : null;
                tmp["bsd_text6"] = item.Contains("bsd_text6") ? item["bsd_text6"] : null;
                tmp["bsd_text7"] = item.Contains("bsd_text7") ? item["bsd_text7"] : null;
                tmp["bsd_text8"] = item.Contains("bsd_text8") ? item["bsd_text8"] : null;
                tmp["bsd_text9"] = item.Contains("bsd_text9") ? item["bsd_text9"] : null;
                tmp["bsd_text10"] = item.Contains("bsd_text10") ? item["bsd_text10"] : null;
            }
        }
    }
}