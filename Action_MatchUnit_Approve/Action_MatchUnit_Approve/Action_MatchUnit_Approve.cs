using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_MatchUnit_Approve
{
    public class Action_MatchUnit_Approve : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private IOrganizationServiceFactory _serviceFactory;

        EntityReference targetEntityRef = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);
            try
            {
                this.targetEntityRef = (EntityReference)_context.InputParameters["Target"];

                UpdateMatchUnit();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void UpdateMatchUnit()
        {
            Entity matchUnit = new Entity(this.targetEntityRef.LogicalName, this.targetEntityRef.Id);
            matchUnit["statuscode"] = new OptionSetValue(100000001); // Approved
            matchUnit["bsd_approver"] = new EntityReference("systemuser", this._context.UserId);
            matchUnit["bsd_approvedate"] = RetrieveLocalTimeFromUTCTime(DateTime.Now, this._service);
            this._service.Update(matchUnit);
        }
        private void UpdateQueue()
        {
            // To do
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
