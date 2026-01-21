using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugin_Booking_Expired_CreateTask
{
    public class Plugin_Booking_Expired_CreateTask : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationServiceFactory serviceFactory = null;
        private IOrganizationService service = null;
        private ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this.context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this.serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory)); 
            this.service = serviceFactory.CreateOrganizationService(this.context.UserId);
            this.tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                if (this.context.Depth > 3) return;
                if (this.context.MessageName != "Update") return;
                if (this.context.InputParameters.Contains("Target") && this.context.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)this.context.InputParameters["Target"];
                    Entity booking = this.service.Retrieve(target.LogicalName, target.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("bsd_expired",
                        "bsd_customerid", "bsd_name"));

                    if (booking.Contains("bsd_expired") && (bool)booking["bsd_expired"] == false)
                        return;
                    // Create a Task when a Booking is expired
                    Entity task = new Entity("task");
                    task["subject"] = "Giữ chỗ \"" + booking["bsd_name"] + "\" của khách hàng \""+ ((EntityReference)booking["bsd_customerid"]).Name + "\" đã hết thời gian";
                    task["bsd_customer"] = new EntityReference(((EntityReference)booking["bsd_customerid"]).LogicalName, ((EntityReference)booking["bsd_customerid"]).Id);
                    task["actualdurationminutes"] = 4320; // 3 days
                    task["scheduledstart"] = RetrieveLocalTimeFromUTCTime(DateTime.Now,this.service);
                    task["regardingobjectid"] = new EntityReference(booking.LogicalName, booking.Id);
                    this.service.Create(task);
                }
            }
            catch (Exception ex)
            {
                this.tracingService.Trace("Plugin_Booking_Expired_CreateTask: {0}", ex.ToString());
                throw;
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
