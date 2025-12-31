using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;
namespace Action_quotatitonReservation_Approve
{
    public class Action_quotatitonReservation_Approve : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            StringBuilder trmess = new StringBuilder();
            try
            {

                ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                if (context.Depth > 3)
                   return;
                //string[] time = ((string)context.InputParameters["strSign"]).Split('/');
                //DateTime dateSign = new DateTime(int.Parse(time[2]), int.Parse(time[1]), int.Parse(time[0]));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity Quote = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(true));
                Entity upquote = new Entity(Quote.LogicalName, Quote.Id);

                upquote["bsd_debtapprovaldate"] = DateTime.Today;
                upquote["bsd_debtapprover"] = new EntityReference("systemuser", context.UserId);
                service.Update(upquote);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
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

