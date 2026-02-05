using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_MatchUnit_Cancel
{
    public class Action_MatchUnit_Cancel : IPlugin
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
                Entity enMatchUnit = this._service.Retrieve(this.targetEntityRef.LogicalName, this.targetEntityRef.Id, new ColumnSet("bsd_queue",
                    "bsd_unit", "bsd_project"));

                DateTime cancelDate = DateTime.Now;
                int priorityNumber = GetPriorityNumber(enMatchUnit);
                UpdateMatchUnit(cancelDate);
                UpdateBooking(enMatchUnit, cancelDate);
                UpdateBookingProjects(enMatchUnit);
                UpdateBookingUnits(enMatchUnit, priorityNumber);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        // Cập nhật thông tin Ráp căn
        private void UpdateMatchUnit(DateTime cancelDate)
        {
            this._tracingService.Trace("Start Update Match Unit");
            Entity matchUnit = new Entity(this.targetEntityRef.LogicalName, this.targetEntityRef.Id);
            matchUnit["statecode"] = new OptionSetValue(1); // inactive
            matchUnit["statuscode"] = new OptionSetValue(100000002); // Cancel
            matchUnit["bsd_canceler"] = new EntityReference("systemuser", this._context.UserId);
            matchUnit["bsd_canceldate"] = RetrieveLocalTimeFromUTCTime(cancelDate, this._service);
            this._service.Update(matchUnit);
            this._tracingService.Trace("End Update Match Unit");
        }
        private void UpdateBooking(Entity enMatchUnit, DateTime cancelDate)
        {
            this._tracingService.Trace("Start update booking");
            Entity enBooking = new Entity(((EntityReference)enMatchUnit["bsd_queue"]).LogicalName, ((EntityReference)enMatchUnit["bsd_queue"]).Id);
            enBooking["bsd_unit"] = null;
            enBooking["bsd_pricelist"] = null;
            enBooking["bsd_usableareasqm"] = null;
            enBooking["bsd_builtupareasqm"] = null;
            enBooking["bsd_usableunitprice"] = null;
            enBooking["bsd_builtupunitprice"] = null;
            enBooking["bsd_price"] = null;
            enBooking["bsd_bookingtime"] = RetrieveLocalTimeFromUTCTime(cancelDate, this._service);
            enBooking["bsd_queuingexpired"] = BookingExpired(enMatchUnit, cancelDate);
            enBooking["bsd_dateorder"] = GetPaymentConfirmDate(enMatchUnit);
            enBooking["statuscode"] = new OptionSetValue(100000004);// Booking
            enBooking["bsd_souutien"] = null;
            enBooking["bsd_douutien"] = 1;
            this._service.Update(enBooking);
            this._tracingService.Trace("End update booking");
        }
        private DateTime BookingExpired(Entity enMatchUnit, DateTime cancelDate)
        {
            this._tracingService.Trace("bookgin expried");
            Entity enProject = this._service.Retrieve(((EntityReference)enMatchUnit["bsd_project"]).LogicalName,
                ((EntityReference)enMatchUnit["bsd_project"]).Id, new ColumnSet("bsd_longqueuingtime"));
            if(enProject.Contains("bsd_longqueuingtime"))
                return RetrieveLocalTimeFromUTCTime(cancelDate, this._service).AddDays((int)enProject["bsd_longqueuingtime"]);
            else
                return RetrieveLocalTimeFromUTCTime(cancelDate, this._service);
        }
        private DateTime GetPaymentConfirmDate(Entity enMatchUnit)
        {
            this._tracingService.Trace("Get Payment Confirm Date");
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_payment"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_confirmeddate"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""statuscode"" operator=""eq"" value=""100000000"" />
                  <condition attribute=""bsd_queue"" operator=""eq"" value=""{((EntityReference)enMatchUnit["bsd_queue"]).Id}"" />
                </filter>
                <order attribute=""createdon"" descending=""true"" />
              </entity>
            </fetch>";
            var result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (result.Entities.Count > 0)
            {
                Entity enPayment = result.Entities[0].ToEntity<Entity>();
                if (enPayment.Contains("bsd_confirmeddate"))
                {
                    DateTime confirmedDate = (DateTime)enPayment["bsd_confirmeddate"];
                    return RetrieveLocalTimeFromUTCTime(confirmedDate, this._service);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Payment does not have Confirmed Date");
                }
            }
            else
            {
                throw new InvalidPluginExecutionException("Can't find confirmed payment for this booking " + ((EntityReference)enMatchUnit["bsd_queue"]).Name);
            }
        }
        private void UpdateBookingProjects(Entity enMatchUnit)
        {
            this._tracingService.Trace("Start update booking project");
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_douutien"" />
                <attribute name=""bsd_opportunityid"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{((EntityReference)enMatchUnit["bsd_queue"]).Id}"" />
                  <condition attribute=""bsd_unit"" operator=""null"" />
                  <condition attribute=""bsd_queueforproject"" operator=""eq"" value=""1"" />
                </filter>
                <order attribute=""createdon"" />
              </entity>
            </fetch>";
            var result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (result.Entities.Count > 0)
            {
                foreach (var entity in result.Entities) {
                    Entity enBookingProject = new Entity(entity.LogicalName, entity.Id);
                    enBookingProject["statuscode"] = new OptionSetValue(100000003);//Waiting
                    enBookingProject["bsd_douutien"] = (int)entity["bsd_douutien"] + 1;
                    this._service.Update(enBookingProject);
                }
            }
        }
        private void UpdateBookingUnits(Entity enMatchUnit,int priorityNumber)
        {
            this._tracingService.Trace("Start update booking units priority");
            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_opportunity"">
                <attribute name=""bsd_name"" />
                <attribute name=""bsd_opportunityid"" />
                <attribute name=""bsd_souutien"" />
                <attribute name=""statuscode"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_opportunityid"" operator=""ne"" value=""{((EntityReference)enMatchUnit["bsd_queue"]).Id}"" />
                  <condition attribute=""bsd_unit"" operator=""eq"" value=""{((EntityReference)enMatchUnit["bsd_unit"]).Id}"" />
                  <condition attribute=""bsd_souutien"" operator=""gt"" value=""{priorityNumber}"" />
                </filter>
                <order attribute=""createdon"" />
              </entity>
            </fetch>";
            var result = this._service.RetrieveMultiple(new FetchExpression(fetchXml));
            foreach (var entity in result.Entities)
            {
                if ((int)entity["bsd_souutien"] > 0)
                {
                    Entity enBookingUnit = new Entity(entity.LogicalName, entity.Id);
                    if ((int)entity["bsd_souutien"] == 2)
                        enBookingUnit["statuscode"] = new OptionSetValue(100000004);//Booking
                    else
                        enBookingUnit["statuscode"] = new OptionSetValue(100000003);//Waiting

                    enBookingUnit["bsd_souutien"] = (int)entity["bsd_souutien"] - 1;
                    this._service.Update(enBookingUnit);
                }
            }
            this._tracingService.Trace("Start update booking units priority");
        }
        private int GetPriorityNumber(Entity enMatchUnit)
        {
            this._tracingService.Trace("Get Priority Number");
            if(!enMatchUnit.Contains("bsd_queue")) return 0;
            Entity enBooking = this._service.Retrieve(((EntityReference)enMatchUnit["bsd_queue"]).LogicalName,
                ((EntityReference)enMatchUnit["bsd_queue"]).Id, new ColumnSet("bsd_souutien"));
            if (enBooking.Contains("bsd_souutien"))
                return (int)enBooking["bsd_souutien"];
            else
                return 0;
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
