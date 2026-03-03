using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace Action_AgingSimulation_GenerateUnit
{
    public class Action_AgingSimulation_GenerateUnit : IPlugin
    {
        public IOrganizationService service = null;
        IOrganizationServiceFactory factory = null;
        ITracingService traceService = null;
        StringBuilder strMess = new StringBuilder();
        StringBuilder strMess2 = new StringBuilder();
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            traceService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = factory.CreateOrganizationService(context.UserId);
            string input01 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input01"]))
            {
                input01 = context.InputParameters["input01"].ToString();
            }
            string input02 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input02"]))
            {
                input02 = context.InputParameters["input02"].ToString();
            }
            string input03 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input03"]))
            {
                input03 = context.InputParameters["input03"].ToString();
            }
            string input04 = "";
            if (!string.IsNullOrEmpty((string)context.InputParameters["input04"]))
            {
                input04 = context.InputParameters["input04"].ToString();
            }
            if (input01 == "Buoc 01" && input02 != "")
            {
                traceService.Trace("Bước 01");
                Entity enTarget = new Entity("bsd_interestsimulation");
                enTarget.Id = Guid.Parse(input02);
                Entity enInterestsimulation = service.Retrieve(enTarget.LogicalName, enTarget.Id, new ColumnSet(new string[4]
                  {
                    "bsd_project",
                    "bsd_block",
                    "bsd_floor",
                    "bsd_floorto"
                  }));
                if (!enInterestsimulation.Contains("bsd_project"))
                    throw new InvalidPluginExecutionException("Please input Project!");
                EntityReference block = enInterestsimulation.Contains("bsd_block") ? (EntityReference)enInterestsimulation["bsd_block"] : null;
                EntityReference floor = enInterestsimulation.Contains("bsd_floor") ? (EntityReference)enInterestsimulation["bsd_floor"] : null;
                EntityReference floorto = enInterestsimulation.Contains("bsd_floorto") ? (EntityReference)enInterestsimulation["bsd_floorto"] : null;
                EntityReference project = (EntityReference)enInterestsimulation["bsd_project"];
                EntityCollection unit = findUnit(project, block, floor, floorto);
                List<string> listUnit = new List<string>();
                foreach (Entity item in unit.Entities)
                {
                    listUnit.Add(item.Id.ToString());
                }
                enTarget["bsd_powerautomate"] = true;
                enTarget["bsd_generate"] = true;
                enTarget["bsd_list"] = string.Join(";", listUnit);
                service.Update(enTarget);
            }
            else if (input01 == "Buoc 02" && input02 != "" && input03 != "" && input04 != "")
            {
                traceService.Trace("Buoc 02");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                EntityReference master = new EntityReference("bsd_interestsimulation", Guid.Parse(input02));
                var fetchXml = $@"
                            <fetch>
                              <entity name='bsd_salesorder'>
                                <attribute name='bsd_name' />
                                <attribute name='bsd_salesorderid' />
                                <attribute name='bsd_unitnumber' />
                                <filter type='and'>
                                  <condition attribute='bsd_unitnumber' operator='eq' value='{input03}'/>
                                  <condition attribute='statuscode' operator='ne' value='100000014'/>
                                  <condition attribute='statuscode' operator='ne' value='100000012'/>
                                  <condition attribute='statuscode' operator='ne' value='2'/>
                                </filter>
                              </entity>
                            </fetch>";
                EntityCollection enOE = service.RetrieveMultiple(new FetchExpression(fetchXml));
                foreach (Entity item in enOE.Entities)
                {
                    Entity enCreate = new Entity("bsd_aginginterestsimulationoption");
                    enCreate["bsd_name"] = item["bsd_name"];
                    enCreate["bsd_aginginterestsimulation"] = master;
                    enCreate["bsd_optionentry"] = item.ToEntityReference();
                    service.Create(enCreate);
                }
            }
            else if (input01 == "Buoc 03" && input02 != "" && input04 != "")
            {
                traceService.Trace("Buoc 03");
                service = factory.CreateOrganizationService(Guid.Parse(input04));
                Entity enConfirmPayment = new Entity("bsd_interestsimulation");
                enConfirmPayment.Id = Guid.Parse(input02);
                enConfirmPayment["bsd_powerautomate"] = false;
                enConfirmPayment["bsd_generate"] = false;
                enConfirmPayment["bsd_errorincalculation"] = "";
                service.Update(enConfirmPayment);
            }
        }
        private EntityCollection findUnit(EntityReference project, EntityReference block, EntityReference floor, EntityReference floorto)
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendLine("<fetch version='1.0' output-format='xml-platform' mapping='logical'>");
            xml.AppendLine("<entity name='bsd_product'>");
            xml.AppendLine("<attribute name='bsd_productid' />");
            xml.AppendLine("<attribute name='statuscode' />");
            xml.AppendLine("<filter type='and'>");
            xml.AppendLine(string.Format("<condition attribute='bsd_projectcode' operator='eq' value='{0}'/>", project.Id));
            xml.AppendLine("<condition attribute='statuscode' operator='in'>");
            xml.AppendLine("<value>100000001</value>");
            xml.AppendLine("<value>100000002</value>");
            xml.AppendLine("</condition>");
            if (block != null && floor != null && floorto == null)
            {
                xml.AppendLine(string.Format("<condition attribute='bsd_blocknumber' operator='eq' value='{0}'/>", block.Id));
                xml.AppendLine(string.Format("<condition attribute='bsd_floor' operator='eq' value='{0}'/>", floor.Id));
            }
            else if (block != null && floor != null && floorto != null)
            {
                int floorNumber1 = toFloorNumber(((DataCollection<string, object>)service.Retrieve(floor.LogicalName, floor.Id, new ColumnSet(true)).Attributes)["bsd_floor"].ToString());
                int floorNumber2 = toFloorNumber(((DataCollection<string, object>)service.Retrieve(floorto.LogicalName, floorto.Id, new ColumnSet(true)).Attributes)["bsd_floor"].ToString());
                EntityCollection floor1 = getFloor(project, block, floorNumber1, floorNumber2);
                xml.AppendLine(string.Format("<condition attribute='bsd_blocknumber' operator='eq' value='{0}'/>", block.Id));
                if (floor1.Entities.Count > 0)
                {
                    xml.AppendLine("<filter type='or'>");
                    foreach (Entity item in floor1.Entities)
                    {
                        xml.AppendLine(string.Format("<condition attribute='bsd_floor' operator='eq' value='{0}'/>", item.Id));
                    }
                    xml.AppendLine("</filter>");
                }
            }
            else if (block != null)
            {
                xml.AppendLine(string.Format("<condition attribute='bsd_blocknumber' operator='eq' value='{0}'/>", block.Id));
            }
            else if (floor != null)
            {
                xml.AppendLine(string.Format("<condition attribute='bsd_floor' operator='eq' value='{0}'/>", floor.Id));
            }
            xml.AppendLine("</filter>");
            xml.AppendLine("</entity>");
            xml.AppendLine("</fetch>");
            EntityCollection unit1 = service.RetrieveMultiple(new FetchExpression(xml.ToString()));
            traceService.Trace("unit1 " + unit1.Entities.Count);
            return unit1;
        }
        private EntityCollection getFloor(EntityReference project, EntityReference block, int from, int to)
        {
            StringBuilder xml = new StringBuilder();
            xml.AppendLine("<fetch version='1.0' output-format='xml-platform' mapping='logical'>");
            xml.AppendLine("<entity name='bsd_floor'>");
            xml.AppendLine("<attribute name='bsd_floor' />");
            xml.AppendLine("<filter type='and'>");
            xml.AppendLine(string.Format("<condition attribute='bsd_project' operator='eq' value='{0}'/>", project.Id)); ;
            xml.AppendLine(string.Format("<condition attribute='bsd_block' operator='eq' value='{0}'/>", block.Id)); ;
            xml.AppendLine("</filter>");
            xml.AppendLine("</entity>");
            xml.AppendLine("</fetch>");
            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(xml.ToString()));
            EntityCollection floor = new EntityCollection();
            foreach (Entity entity in (Collection<Entity>)entityCollection.Entities)
            {
                int floorNumber = toFloorNumber(((DataCollection<string, object>)entity.Attributes)["bsd_floor"].ToString());
                if (floorNumber >= from && floorNumber <= to)
                    floor.Entities.Add(entity);
            }
            return floor;
        }
        private int toFloorNumber(string floor)
        {
            string upper = floor.ToUpper();
            byte[] bytes = Encoding.ASCII.GetBytes(upper);
            string str1 = "";
            string str2 = "";
            for (int index = 0; index < upper.Length; ++index)
            {
                if (bytes[index] >= (byte)48 && bytes[index] <= (byte)57)
                    str1 += upper[index].ToString();
                else
                    str2 += upper[index].ToString();
            }
            switch (Convert.ToInt32(str1).ToString() + str2)
            {
                case "3A":
                    return 4;
                case "12A":
                    return 13;
                case "12B":
                    return 14;
                default:
                    return Convert.ToInt32(str1);
            }
        }
    }
}
