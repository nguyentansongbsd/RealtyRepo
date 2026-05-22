using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_CollectionMeeting_Close
{
    public class Action_CollectionMeeting_Close : IPlugin
    {
        IOrganizationService service = null;
        ITracingService traceService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                traceService.Trace("start");
                if (context.Depth > 1) return;

                EntityReference target = (EntityReference)context.InputParameters["Target"];
                Entity enAppointment = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "regardingobjectid" }));
                if (!enAppointment.Contains("regardingobjectid"))
                    return;

                EntityReference refRO = (EntityReference)enAppointment["regardingobjectid"];
                if (!"bsd_followuplist".Equals(refRO.LogicalName))
                    return;

                traceService.Trace("1");
                Entity enFUL = service.Retrieve(refRO.LogicalName, refRO.Id, new ColumnSet(true));
                bool isTermination = enFUL.Contains("bsd_termination") ? (bool)enFUL["bsd_termination"] : false;
                bool isTerminateLetter = enFUL.Contains("bsd_terminateletter") ? (bool)enFUL["bsd_terminateletter"] : false;
                EntityReference refUnit = (EntityReference)enFUL["bsd_units"];

                UpdateCollectionMeeting(enAppointment);
                UpdateFUL(enFUL);
                UpdateUnit(refUnit);

                if (isTermination && isTerminateLetter || isTermination)
                {
                    CreateTermination(enFUL, refUnit);
                }
                else if (isTerminateLetter)
                {
                    CreateTerminationLetter(enFUL, refUnit);
                }

                traceService.Trace("done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void CreateTermination(Entity enFUL, EntityReference refUnit)
        {
            traceService.Trace("CreateTermination");

            Entity newTer = new Entity("bsd_termination");
            newTer["bsd_name"] = $"Termination Letter {refUnit.Name}";
            newTer["bsd_terminationdate"] = DateTime.UtcNow;
            newTer["bsd_terminationtype"] = ValidValue(enFUL, "bsd_terminationtype");
            newTer["bsd_resell"] = ValidValue(enFUL, "bsd_resell");

            newTer["bsd_project"] = ValidValue(enFUL, "bsd_project");
            newTer["bsd_units"] = ValidValue(enFUL, "bsd_units");
            newTer["bsd_source"] = ValidValue(enFUL, "bsd_source");
            newTer["bsd_reservation"] = ValidValue(enFUL, "bsd_reservation");
            newTer["bsd_reservationcontract"] = ValidValue(enFUL, "bsd_reservationcontract");
            newTer["bsd_optionentry"] = ValidValue(enFUL, "bsd_optionentry");

            newTer["bsd_typeforfeiture"] = ValidValue(enFUL, "bsd_typeforfeiture");
            newTer["bsd_customer"] = ValidValue(enFUL, "bsd_customer");
            newTer["bsd_penaltytype"] = ValidValue(enFUL, "bsd_penaltytype");
            newTer["bsd_forfeiturepercent"] = ValidValue(enFUL, "bsd_forfeiturepercent");
            newTer["bsd_penaltybase"] = ValidValue(enFUL, "bsd_penaltybase");
            newTer["bsd_amount"] = ValidValue(enFUL, "bsd_amount");
            newTer["bsd_otherpenalty"] = ValidValue(enFUL, "bsd_otherpenalty");
            newTer["bsd_totalfines"] = ValidValue(enFUL, "bsd_totalfines");

            newTer["bsd_developerpenalty"] = ValidValue(enFUL, "bsd_developerpenalty");
            newTer["bsd_developer"] = ValidValue(enFUL, "bsd_developer");
            newTer["bsd_penaltytypedeveloper"] = ValidValue(enFUL, "bsd_penaltytypedeveloper");
            newTer["bsd_forfeiturepercentdeveloper"] = ValidValue(enFUL, "bsd_forfeiturepercentdeveloper");
            newTer["bsd_penaltybasedeveloper"] = ValidValue(enFUL, "bsd_penaltybasedeveloper");
            newTer["bsd_amountdeveloper"] = ValidValue(enFUL, "bsd_amountdeveloper");
            newTer["bsd_otherpenaltydeveloper"] = ValidValue(enFUL, "bsd_otherpenaltydeveloper");
            newTer["bsd_totalfinesdeveloper"] = ValidValue(enFUL, "bsd_totalfinesdeveloper");

            newTer["bsd_depositfee"] = ValidValue(enFUL, "bsd_depositfee");
            newTer["bsd_totalamountpaid"] = ValidValue(enFUL, "bsd_totalamountpaid");

            newTer.Id = Guid.NewGuid();
            service.Create(newTer);
        }

        private object ValidValue(Entity enFUL, string field)
        {
            return enFUL.Contains(field) ? enFUL[field] : null;
        }

        private void CreateTerminationLetter(Entity enFUL, EntityReference refUnit)
        {
            traceService.Trace("CreateTerminationLetter");

            Entity newTerminationLetter = new Entity("bsd_terminateletter");
            newTerminationLetter["bsd_name"] = $"Termination Letter {refUnit.Name}";
            newTerminationLetter["bsd_customer"] = ValidValue(enFUL, "bsd_customer");
            newTerminationLetter["bsd_date"] = DateTime.UtcNow;
            newTerminationLetter["bsd_source"] = ValidValue(enFUL, "bsd_source");
            newTerminationLetter["bsd_system"] = true;

            newTerminationLetter["bsd_project"] = ValidValue(enFUL, "bsd_project");
            newTerminationLetter["bsd_units"] = ValidValue(enFUL, "bsd_units");
            newTerminationLetter["bsd_reservation"] = ValidValue(enFUL, "bsd_reservation");
            newTerminationLetter["bsd_reservationcontract"] = ValidValue(enFUL, "bsd_reservationcontract");
            newTerminationLetter["bsd_optionentry"] = ValidValue(enFUL, "bsd_optionentry");

            newTerminationLetter["bsd_followuplist"] = enFUL.ToEntityReference();
            newTerminationLetter["bsd_totalforfeitureamount"] = ValidValue(enFUL, "bsd_totalforfeitureamount");

            newTerminationLetter.Id = Guid.NewGuid();
            service.Create(newTerminationLetter);
        }

        private void UpdateCollectionMeeting(Entity enAppointment)
        {
            traceService.Trace("UpdateCollectionMeeting");

            Entity upCM = new Entity(enAppointment.LogicalName, enAppointment.Id);
            upCM["statecode"] = new OptionSetValue(1); //Completed
            service.Update(upCM);
        }

        private void UpdateFUL(Entity enFUL)
        {
            traceService.Trace("UpdateFUL");

            Entity upFUL = new Entity(enFUL.LogicalName, enFUL.Id);
            upFUL["statuscode"] = new OptionSetValue(100000000); //Completed
            service.Update(upFUL);
        }

        private void UpdateUnit(EntityReference refUnit)
        {
            traceService.Trace("UpdateUnit");

            Entity upUnit = new Entity(refUnit.LogicalName, refUnit.Id);
            upUnit["bsd_isfollowuplist"] = false;
            service.Update(upUnit);
        }
    }
}