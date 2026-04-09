using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Lead_Merge
{
    public class Action_Lead_Merge : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;

        public class Output
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("start");

            string data = context.InputParameters.Contains("data") && !string.IsNullOrWhiteSpace((string)context.InputParameters["data"]) ? context.InputParameters["data"].ToString() : string.Empty;
            if (!string.IsNullOrWhiteSpace(data))
            {
                List<Output> listData = JsonConvert.DeserializeObject<List<Output>>(data);
                var result = listData.ToDictionary(item => item.Name, item => item.Value);

                Entity upLead = new Entity("bsd_lead", new Guid(result["radio-primary"]));
                traceService.Trace("1");
                if (!string.IsNullOrWhiteSpace(result["radio-source"])) upLead["bsd_leadsourcecode"] = new OptionSetValue(int.Parse(result["radio-source"]));
                if (!string.IsNullOrWhiteSpace(result["radio-rating"])) upLead["bsd_leadqualitycode"] = new OptionSetValue(int.Parse(result["radio-rating"]));
                if (!string.IsNullOrWhiteSpace(result["radio-status"])) upLead["statuscode"] = new OptionSetValue(int.Parse(result["radio-status"]));

                traceService.Trace("2");
                upLead["bsd_subject"] = result["radio-topic"];
                upLead["bsd_firstname"] = result["radio-first-name"];
                upLead["bsd_lastname"] = result["radio-last-name"];
                upLead["bsd_jobtitle"] = result["radio-job-title"];
                upLead["bsd_telephone1"] = result["radio-business-phone"];
                upLead["bsd_mobilephone"] = result["radio-mobile-phone"];
                upLead["bsd_emailaddress1"] = result["radio-email"];

                traceService.Trace("3");
                upLead["bsd_companyname"] = result["radio-company-name"];
                upLead["bsd_websiteurl"] = result["radio-website"];
                upLead["bsd_address1_line1"] = result["radio-street-1"];
                upLead["bsd_address1_line2"] = result["radio-street-2"];
                upLead["bsd_address1_line3"] = result["radio-street-3"];
                upLead["bsd_address1_city"] = result["radio-city"];
                upLead["bsd_address1_stateorprovince"] = result["radio-state"];
                upLead["bsd_address1_postalcode"] = result["radio-zip"];
                upLead["bsd_address1_country"] = result["radio-country"];

                traceService.Trace("4");
                upLead["bsd_description"] = result["radio-description"];
                if (!string.IsNullOrWhiteSpace(result["radio-industry"])) upLead["bsd_industrycode"] = new OptionSetValue(int.Parse(result["radio-industry"]));
                if (!string.IsNullOrWhiteSpace(result["radio-annual-revenue"])) upLead["bsd_revenue"] = new Money(decimal.Parse(result["radio-annual-revenue"]));
                if (!string.IsNullOrWhiteSpace(result["radio-eployees"])) upLead["bsd_numberofemployees"] = int.Parse(result["radio-eployees"]);
                upLead["bsd_sic"] = result["radio-sic"];
                if (!string.IsNullOrWhiteSpace(result["radio-currency"])) upLead["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid(result["radio-currency"]));

                traceService.Trace("5");
                if (!string.IsNullOrWhiteSpace(result["radio-preferred"])) upLead["bsd_preferredcontactmethodcode"] = new OptionSetValue(int.Parse(result["radio-preferred"]));
                upLead["bsd_donotemail"] = bool.Parse(result["radio-donotemail"]);
                upLead["bsd_followemail"] = bool.Parse(result["radio-follow-email"]);
                upLead["bsd_donotbulkemail"] = bool.Parse(result["radio-bulk-email"]);
                upLead["bsd_donotphone"] = bool.Parse(result["radio-phone"]);
                upLead["bsd_donotpostalmail"] = bool.Parse(result["radio-donotpostalmail"]);

                service.Update(upLead);
                traceService.Trace("6");

                Entity upLeadSecond = new Entity("bsd_lead", new Guid(result["radio-primary-second"]));
                upLeadSecond["statuscode"] = new OptionSetValue(100000002);
                service.Update(upLeadSecond);
                traceService.Trace("7");
            }
        }
    }
}