using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_Lead_Qualify
{
    public class Action_Lead_Qualify : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            traceService.Trace("start");

            EntityReference refLead = (EntityReference)context.InputParameters["Target"];
            Entity enLead = service.Retrieve(refLead.LogicalName, refLead.Id, new ColumnSet(true));

            Guid idContact = Guid.Empty;
            Guid idAccount = Guid.Empty;
            if (enLead.Contains("bsd_firstname") || enLead.Contains("bsd_lastname"))
            {
                idContact = CreateContact(enLead, refLead);
            }

            if (enLead.Contains("bsd_companyname"))
            {
                idAccount = CreateAccount(enLead, refLead, idContact);

                Entity upContact = new Entity("contact", idContact);
                upContact["parentcustomerid"] = new EntityReference("account", idAccount);
                service.Update(upContact);
            }

            Guid idOpportunity = CreateOpportunity(enLead, refLead, idContact, idAccount);

            Entity upLead = new Entity(enLead.LogicalName, enLead.Id);
            upLead["statuscode"] = new OptionSetValue(100000001);
            if (idContact != Guid.Empty) upLead["bsd_parentcontact"] = new EntityReference("contact", idContact);
            if (idAccount != Guid.Empty) upLead["bsd_parentaccount"] = new EntityReference("account", idAccount);
            service.Update(upLead);

            context.OutputParameters["id"] = idOpportunity.ToString();
        }

        private Guid CreateContact(Entity enLead, EntityReference refLead)
        {
            traceService.Trace("CreateContact");
            Entity newContact = new Entity("contact");
            newContact["firstname"] = enLead.Contains("bsd_firstname") ? enLead["bsd_firstname"] : null;
            newContact["lastname"] = enLead.Contains("bsd_lastname") ? enLead["bsd_lastname"] : null;
            newContact["jobtitle"] = enLead.Contains("bsd_jobtitle") ? enLead["bsd_jobtitle"] : null;
            newContact["telephone1"] = enLead.Contains("bsd_telephone1") ? enLead["bsd_telephone1"] : null;
            newContact["mobilephone"] = enLead.Contains("bsd_mobilephone") ? enLead["bsd_mobilephone"] : null;
            newContact["emailaddress1"] = enLead.Contains("bsd_emailaddress1") ? enLead["bsd_emailaddress1"] : null;
            newContact["websiteurl"] = enLead.Contains("bsd_websiteurl") ? enLead["bsd_websiteurl"] : null;

            newContact["address1_line1"] = enLead.Contains("bsd_address1_line1") ? enLead["bsd_address1_line1"] : null;
            newContact["address1_line2"] = enLead.Contains("bsd_address1_line2") ? enLead["bsd_address1_line2"] : null;
            newContact["address1_line3"] = enLead.Contains("bsd_address1_line3") ? enLead["bsd_address1_line3"] : null;
            newContact["address1_city"] = enLead.Contains("bsd_address1_city") ? enLead["bsd_address1_city"] : null;
            newContact["address1_stateorprovince"] = enLead.Contains("bsd_address1_stateorprovince") ? enLead["bsd_address1_stateorprovince"] : null;
            newContact["address1_postalcode"] = enLead.Contains("bsd_address1_postalcode") ? enLead["bsd_address1_postalcode"] : null;
            newContact["address1_country"] = enLead.Contains("bsd_address1_country") ? enLead["bsd_address1_country"] : null;

            newContact["bsd_originatinglead"] = refLead;
            newContact["description"] = enLead.Contains("bsd_description") ? enLead["bsd_description"] : null;

            if (enLead.Contains("bsd_preferredcontactmethodcode"))
            {
                switch (((OptionSetValue)enLead["bsd_preferredcontactmethodcode"]).Value)
                {
                    case 100000000:
                        newContact["preferredcontactmethodcode"] = new OptionSetValue(1);
                        break;
                    case 100000001:
                        newContact["preferredcontactmethodcode"] = new OptionSetValue(2);
                        break;
                    case 100000002:
                        newContact["preferredcontactmethodcode"] = new OptionSetValue(3);
                        break;
                    case 100000003:
                        newContact["preferredcontactmethodcode"] = new OptionSetValue(4);
                        break;
                    case 100000004:
                        newContact["preferredcontactmethodcode"] = new OptionSetValue(5);
                        break;
                }
            }
            newContact["donotemail"] = enLead.Contains("bsd_donotemail") ? enLead["bsd_donotemail"] : null;
            newContact["followemail"] = enLead.Contains("bsd_followemail") ? enLead["bsd_followemail"] : null;
            newContact["donotbulkemail"] = enLead.Contains("bsd_donotbulkemail") ? enLead["bsd_donotbulkemail"] : null;
            newContact["donotphone"] = enLead.Contains("bsd_donotphone") ? enLead["bsd_donotphone"] : null;
            newContact["donotpostalmail"] = enLead.Contains("bsd_donotpostalmail") ? enLead["bsd_donotpostalmail"] : null;

            newContact.Id = Guid.NewGuid();
            Guid id = service.Create(newContact);
            return id;
        }

        private Guid CreateAccount(Entity enLead, EntityReference refLead, Guid idContact)
        {
            traceService.Trace("CreateAccount");
            Entity newAccount = new Entity("account");
            newAccount["name"] = enLead["bsd_companyname"];
            if (idContact != Guid.Empty) newAccount["primarycontactid"] = new EntityReference("contact", idContact);
            newAccount["telephone1"] = enLead.Contains("bsd_telephone1") ? enLead["bsd_telephone1"] : null;
            newAccount["emailaddress1"] = enLead.Contains("bsd_emailaddress1") ? enLead["bsd_emailaddress1"] : null;
            newAccount["websiteurl"] = enLead.Contains("bsd_websiteurl") ? enLead["bsd_websiteurl"] : null;

            newAccount["address1_line1"] = enLead.Contains("bsd_address1_line1") ? enLead["bsd_address1_line1"] : null;
            newAccount["address1_line2"] = enLead.Contains("bsd_address1_line2") ? enLead["bsd_address1_line2"] : null;
            newAccount["address1_line3"] = enLead.Contains("bsd_address1_line3") ? enLead["bsd_address1_line3"] : null;
            newAccount["address1_city"] = enLead.Contains("bsd_address1_city") ? enLead["bsd_address1_city"] : null;
            newAccount["address1_stateorprovince"] = enLead.Contains("bsd_address1_stateorprovince") ? enLead["bsd_address1_stateorprovince"] : null;
            newAccount["address1_postalcode"] = enLead.Contains("bsd_address1_postalcode") ? enLead["bsd_address1_postalcode"] : null;
            newAccount["address1_country"] = enLead.Contains("bsd_address1_country") ? enLead["bsd_address1_country"] : null;

            newAccount["bsd_originatinglead"] = refLead;
            //newAccount["industrycode"] = enLead.Contains("bsd_industrycode") ? enLead["bsd_industrycode"] : null;
            newAccount["sic"] = enLead.Contains("bsd_sic") ? enLead["bsd_sic"] : null;
            newAccount["revenue"] = enLead.Contains("bsd_revenue") ? enLead["bsd_revenue"] : null;
            newAccount["numberofemployees"] = enLead.Contains("bsd_numberofemployees") ? enLead["bsd_numberofemployees"] : null;
            newAccount["description"] = enLead.Contains("bsd_description") ? enLead["bsd_description"] : null;

            if (enLead.Contains("bsd_preferredcontactmethodcode"))
            {
                switch (((OptionSetValue)enLead["bsd_preferredcontactmethodcode"]).Value)
                {
                    case 100000000:
                        newAccount["preferredcontactmethodcode"] = new OptionSetValue(1);
                        break;
                    case 100000001:
                        newAccount["preferredcontactmethodcode"] = new OptionSetValue(2);
                        break;
                    case 100000002:
                        newAccount["preferredcontactmethodcode"] = new OptionSetValue(3);
                        break;
                    case 100000003:
                        newAccount["preferredcontactmethodcode"] = new OptionSetValue(4);
                        break;
                    case 100000004:
                        newAccount["preferredcontactmethodcode"] = new OptionSetValue(5);
                        break;
                }
            }
            newAccount["donotemail"] = enLead.Contains("bsd_donotemail") ? enLead["bsd_donotemail"] : null;
            newAccount["followemail"] = enLead.Contains("bsd_followemail") ? enLead["bsd_followemail"] : null;
            newAccount["donotbulkemail"] = enLead.Contains("bsd_donotbulkemail") ? enLead["bsd_donotbulkemail"] : null;
            newAccount["donotphone"] = enLead.Contains("bsd_donotphone") ? enLead["bsd_donotphone"] : null;
            newAccount["donotpostalmail"] = enLead.Contains("bsd_donotpostalmail") ? enLead["bsd_donotpostalmail"] : null;

            newAccount.Id = Guid.NewGuid();
            Guid id = service.Create(newAccount);
            return id;
        }

        private Guid CreateOpportunity(Entity enLead, EntityReference refLead, Guid idContact, Guid idAccount)
        {
            traceService.Trace("CreateOpportunity");
            Entity newOp = new Entity("bsd_opportunity");
            newOp["bsd_name"] = enLead.Contains("bsd_subject") ? enLead["bsd_subject"] : null;
            newOp["bsd_lead"] = refLead;
            if (idContact != Guid.Empty) newOp["bsd_parentcontact"] = new EntityReference("contact", idContact);
            if (idAccount != Guid.Empty) newOp["bsd_parentaccount"] = new EntityReference("account", idAccount);
            newOp["bsd_description"] = enLead.Contains("bsd_description") ? enLead["bsd_description"] : null;

            newOp.Id = Guid.NewGuid();
            Guid id = service.Create(newOp);
            return id;
        }
    }
}