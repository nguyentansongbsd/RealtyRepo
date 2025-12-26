using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_PaymentScheme_CopyKhacDA
{
    public class Action_PaymentScheme_CopyKhacDA : IPlugin
    {
        IPluginExecutionContext context = null;
        IOrganizationServiceFactory serviceFactory = null;
        IOrganizationService service = null;
        ITracingService tracingService = null;
        EntityReference target = null;
        Entity enPaymentScheme = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            target = (EntityReference)context.InputParameters["Target"];
            
            try
            {
                Init().Wait();
            }
            catch (AggregateException ex)
            {
                var innerEx = ex.InnerExceptions.FirstOrDefault();
                if(innerEx != null)
                {
                    throw innerEx;
                }
            }
        }
        private async Task Init()
        {
            try
            {
                string projectId = (string)context.InputParameters["projectId"];
                enPaymentScheme = this.service.Retrieve(this.target.LogicalName, this.target.Id, new ColumnSet(true));
                await createPaymentSchemeCopy(enPaymentScheme, projectId);
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        private async Task createPaymentSchemeCopy(Entity _enPaymentScheme, string projectId)
        {
            try
            {
                tracingService.Trace("Start copy");
                _enPaymentScheme.Attributes.Remove("bsd_paymentschemeid");
                _enPaymentScheme.Attributes.Remove("bsd_paymentschemecodenew");
                _enPaymentScheme.Attributes.Remove("bsd_comfirmeddate");
                _enPaymentScheme.Attributes.Remove("bsd_comfirmperson");
                _enPaymentScheme.Attributes.Remove("ownerid");

                Guid id = Guid.NewGuid();
                _enPaymentScheme.Id = id;
                _enPaymentScheme["bsd_name"] = (string)this.enPaymentScheme["bsd_name"] + " Copy";
                _enPaymentScheme["bsd_project"] = new EntityReference("bsd_project", Guid.Parse(projectId));
                _enPaymentScheme["statuscode"] = new OptionSetValue(1);

                this.service.Create(_enPaymentScheme);
                copyInstallments(id, projectId);
                this.context.OutputParameters["paymentSchemeId"] = id;
                tracingService.Trace("End copy");
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
        private async Task copyInstallments(Guid paymentSchemeIdNew, string projectId)
        {
            try
            {
                tracingService.Trace("Start copy installment");
                string fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='bsd_paymentschemedetailmaster'>
                    <order attribute='bsd_ordernumber' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='bsd_paymentscheme' operator='eq' value='{this.target.Id}' />
                    </filter>
                  </entity>
                </fetch>";
                var result = this.service.RetrieveMultiple(new FetchExpression(fetchxml));
                if (result == null || result.Entities.Count == 0) return;
                Guid paymentSchemeDetailId = Guid.Empty;
                foreach (var item in result.Entities)
                {
                    Entity enPaymentSchemeDetailNew = item;
                    enPaymentSchemeDetailNew.Attributes.Remove("bsd_paymentschemedetailmasterid");
                    enPaymentSchemeDetailNew.Attributes.Remove("ownerid");

                    enPaymentSchemeDetailNew.Id = Guid.NewGuid();
                    enPaymentSchemeDetailNew["bsd_paymentscheme"] = new EntityReference("bsd_paymentscheme", paymentSchemeIdNew);
                    enPaymentSchemeDetailNew["bsd_project"] = new EntityReference("bsd_project", Guid.Parse(projectId));

                    paymentSchemeDetailId = this.service.Create(enPaymentSchemeDetailNew);
                }
                tracingService.Trace("End copy installment");
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw ex;
            }
        }
    }
}
