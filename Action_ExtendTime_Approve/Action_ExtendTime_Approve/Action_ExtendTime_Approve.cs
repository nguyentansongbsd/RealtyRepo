using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_ExtendTime_Approve
{
    public class Action_ExtendTime_Approve : IPlugin
    {
        private IPluginExecutionContext context = null;
        private IOrganizationServiceFactory factory = null;
        private IOrganizationService service = null;
        private ITracingService tracingService = null;

        private Entity enExtendTime = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this.context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this.factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this.service = factory.CreateOrganizationService(this.context.UserId);
            this.tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                EntityReference enfExtendTime = (EntityReference)this.context.InputParameters["Target"];
                this.enExtendTime = this.service.Retrieve(enfExtendTime.LogicalName, enfExtendTime.Id, new ColumnSet("bsd_extenddate", 
                    "bsd_booking"));

                UpdateExtendTime();
                UpdateBookingExtendDate();
            }
            catch (Exception ex)
            {
                tracingService.Trace("Action_ExtendTime_Approve: {0}", ex.ToString());
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void UpdateExtendTime()
        {
            try
            {
                tracingService.Trace("Start Update Extend Time");
                Entity enExtendTime = new Entity(this.enExtendTime.LogicalName, this.enExtendTime.Id);
                enExtendTime["bsd_approver"] = new EntityReference("systemuser", this.context.UserId);
                enExtendTime["bsd_approvedate"] = RetrieveLocalTimeFromUTCTime(DateTime.UtcNow, this.service);
                enExtendTime["statuscode"] = new OptionSetValue(100000000); // Approved
                this.service.Update(enExtendTime);
                tracingService.Trace("End Update Extend Time");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void UpdateBookingExtendDate()
        {
            try
            {
                tracingService.Trace("Start update Booking");
                
                Entity enBooking = new Entity(((EntityReference)this.enExtendTime["bsd_booking"]).LogicalName, ((EntityReference)this.enExtendTime["bsd_booking"]).Id);
                if(this.enExtendTime.Contains("bsd_extenddate"))
                {
                    DateTime extendDate = (DateTime)this.enExtendTime["bsd_extenddate"];
                    enBooking["bsd_queuingexpired"] = RetrieveLocalTimeFromUTCTime(extendDate, service);
                }
                enBooking["bsd_expired"] = false;
                this.service.Update(enBooking);
                tracingService.Trace("End update Booking");
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
    }
}
