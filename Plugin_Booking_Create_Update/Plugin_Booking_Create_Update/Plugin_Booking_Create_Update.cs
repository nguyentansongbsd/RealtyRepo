using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
namespace Plugin_Booking_Create_Update
{
    public class Plugin_Booking_Create_Update : IPlugin
    {
        IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService trace = null;
        IPluginExecutionContext context = null;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.Depth > 1) return;
            if (context.MessageName == "Create" || context.MessageName == "Update")
            {
                Entity target = (Entity)context.InputParameters["Target"];

                if (context.MessageName == "Update" && target.Contains("bsd_customerid"))
                {
                    EntityReference customer = (EntityReference)target["bsd_customerid"];
                    List<string> missingFields = new List<string>();

                    // Define field mapping
                    Dictionary<string, string> fieldMap = new Dictionary<string, string>();

                    if (customer.LogicalName == "contact")
                    {
                        fieldMap = new Dictionary<string, string>
                            {
                                {"bsd_localization", "Nationality"},
                                {"birthdate", "Birthday"},
                                {"bsd_identitycardnumber", "Identity Card Number (ID)"},
                                {"bsd_country", "Country"},
                                {"bsd_province", "Province"},
                                {"bsd_contactaddress", "Contact Address (VN)"},
                                {"bsd_permanentcountry", "Permanent Country"},
                                {"bsd_permanentprovince", "Permanent Province"},
                                {"bsd_permanentaddress1", "Permanent Address (VN)"}
                            };
                    }
                    else if (customer.LogicalName == "account")
                    {
                        fieldMap = new Dictionary<string, string>
                            {
                                {"bsd_localization", "Nationality"},
                                {"bsd_registrationcode", "Registration Code"},
                                {"bsd_nation", "Country"},
                                {"bsd_province", "Province"},
                                {"bsd_addressvn", "Address (VN)"},
                                {"bsd_permanentcountry", "Permanent Country"},
                                {"bsd_permanentprovince", "Permanent Province"},
                                {"bsd_permanentaddress1", "Permanent Address (VN)"}
                            };
                    }

                    // Retrieve only needed columns
                    Entity customerEntity = service.Retrieve(
                        customer.LogicalName,
                        customer.Id,
                        new ColumnSet(fieldMap.Keys.ToArray())
                    );

                    // Validate null / missing
                    foreach (var field in fieldMap)
                    {
                        if (!customerEntity.Contains(field.Key) || customerEntity[field.Key] == null)
                        {
                            missingFields.Add(field.Value);
                        }
                    }

                    // Throw error if needed
                    if (missingFields.Count > 0)
                    {
                        throw new InvalidPluginExecutionException(
                            "Please fill in the missing customer information below:\r\n ["
                            + string.Join("\r\n| ", missingFields) + "]"
                        );
                    }
                }
            }
        }
    }
}