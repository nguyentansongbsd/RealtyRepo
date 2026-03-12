using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Plugin_WaiverApproval_Create_Update
{
    public class Plugin_WaiverApproval_Create_Update : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            ITracingService traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.MessageName == "Create")
            {

            }
            else if (context.MessageName == "Update")
            {
                if (context.Depth > 2) return;
                Entity target = (Entity)context.InputParameters["Target"];
                Entity enTarget = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                int statuscode = enTarget.Contains("statuscode") ? ((OptionSetValue)enTarget["statuscode"]).Value : 0;
                if (statuscode == 100000000 || statuscode == 100000001)
                {
                    if (!enTarget.Contains("bsd_transactiontype"))
                    {
                        throw new InvalidPluginExecutionException("No Transaction Type value");
                    }
                    int transactiontype = ((OptionSetValue)enTarget["bsd_transactiontype"]).Value;
                    if (transactiontype == 667980000 && !enTarget.Contains("bsd_ra"))
                        throw new InvalidPluginExecutionException("No RA value");
                    else if (transactiontype == 667980001 && !enTarget.Contains("bsd_optionentry"))
                        throw new InvalidPluginExecutionException("No SPA value");
                    if (statuscode == 100000000)//Approve
                    {
                        var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_waiverapprovaldetail"">
                            <attribute name=""bsd_waiverapprovaldetailid"" />
                            <attribute name=""bsd_installment"" />
                            <attribute name=""bsd_waiveramount"" />
                            <attribute name=""bsd_interestchargeamount"" />
                            <attribute name=""bsd_interestwaspaid"" />
                            <attribute name=""bsd_waiverinterest"" />
                            <filter>
                              <condition attribute=""statuscode"" operator=""eq"" value=""{1}"" />
                              <condition attribute=""bsd_waiveramount"" operator=""gt"" value=""0"" />
                              <condition attribute=""bsd_installment"" operator=""not-null"" />
                              <condition attribute=""bsd_waiverapproval"" operator=""eq"" value=""{target.Id}"" />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection enDetail = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (enDetail.Entities.Count == 0) throw new InvalidPluginExecutionException("Waiver approval detail not found.");
                        foreach (Entity entity in enDetail.Entities)
                        {
                            Interest_Charge(entity, (transactiontype == 667980000 ? (EntityReference)enTarget["bsd_ra"] : (EntityReference)enTarget["bsd_optionentry"]), ((Money)entity["bsd_waiveramount"]).Value);
                        }
                    }
                    else if (statuscode == 100000001)//Revert
                    {
                        var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
                        <fetch>
                          <entity name=""bsd_waiverapprovaldetail"">
                            <attribute name=""bsd_waiverapprovaldetailid"" />
                            <attribute name=""bsd_installment"" />
                            <attribute name=""bsd_waiveramount"" />
                            <attribute name=""bsd_interestchargeamount"" />
                            <attribute name=""bsd_interestwaspaid"" />
                            <attribute name=""bsd_waiverinterest"" />
                            <filter>
                              <condition attribute=""statuscode"" operator=""eq"" value=""{100000000}"" />
                              <condition attribute=""bsd_waiveramount"" operator=""gt"" value=""0"" />
                              <condition attribute=""bsd_installment"" operator=""not-null"" />
                              <condition attribute=""bsd_waiverapproval"" operator=""eq"" value=""{target.Id}"" />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection enDetail = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (enDetail.Entities.Count == 0) throw new InvalidPluginExecutionException("Waiver approval detail not found.");
                    }
                }
                else return;
            }
        }
        private void Interest_Charge(Entity enDetail, EntityReference enrHD, decimal amountPay)
        {
            Entity enHD = service.Retrieve(enrHD.LogicalName, enrHD.Id,
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
            service.Update(upHD);
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
                  <condition attribute=""bsd_paymentschemedetailid"" operator=""eq"" value=""{((EntityReference)enDetail["bsd_installment"]).Id}"" />
                </filter>
              </entity>
            </fetch>";
            EntityCollection enIntallment = service.RetrieveMultiple(new FetchExpression(fetchXml));
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
                    upIntallment["bsd_interestwaspaid"] = new Money(bsd_interestwaspaid);
                    upIntallment["bsd_waiverinterest"] = new Money(bsd_waiverinterest + amountPay);
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
                    upIntallment["bsd_interestwaspaid"] = new Money(bsd_interestwaspaid);
                    upIntallment["bsd_waiverinterest"] = new Money(bsd_waiverinterest + bsd_balance);
                    upIntallment["bsd_interestchargeremaining"] = new Money(0);
                    if (bsd_balanceIns == 0)
                    {
                        upIntallment["statuscode"] = new OptionSetValue(100000001);
                        upIntallment["bsd_interestchargestatus"] = new OptionSetValue(100000001);
                    }
                    amountPay -= bsd_balance;
                }
                service.Update(upIntallment);
                if (amountPay <= 0) break;
            }
            if (amountPay > 0)
            {
                throw new InvalidPluginExecutionException("The amount payable is more than the interest charge required.");
            }
            Entity enUp = new Entity(enDetail.LogicalName, enDetail.Id);
            enUp["statuscode"] = new OptionSetValue(100000000);
            service.Update(enUp);
        }
    }
}
