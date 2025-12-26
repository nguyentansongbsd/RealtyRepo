using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace SaleDirectAction
{
    public class SaleDirectActionBackup : IPlugin
    {
        IOrganizationService service;
        IOrganizationServiceFactory factory;

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            EntityReference productRef = (EntityReference)context.InputParameters["Target"];
            string command = context.InputParameters["Command"].ToString();
            if (command == "Book")
            {
                #region Book
                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                string content = "";
                Entity product = RetrieveValidUnit(productRef.Id);
               
                //service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet(
                //new string[] { "name", "statuscode", "defaultuomid", "bsd_listprice", "statecode", "bsd_projectcode", "bsd_blocknumber", "bsd_floor", "bsd_phaseslaunchid" }));
                if (((OptionSetValue)product["statecode"]).Value == 1)
                    throw new Exception("This unit is not public!");
                else if (((OptionSetValue)product["statuscode"]).Value == 100000002)
                    throw new Exception("This unit was sold!");
                else if (!product.Attributes.Contains("bsd_floor"))
                    throw new Exception("Please select floor for this unit!");
                else if (!product.Attributes.Contains("bsd_blocknumber"))
                    throw new Exception("Please select block for this unit!");
                else if (!product.Attributes.Contains("bsd_projectcode"))
                    throw new Exception("Please select project for this unit!");
                else if (!product.Attributes.Contains("defaultuomid"))
                    throw new Exception("Please select default unit for this unit!");
                else
                {

                    Entity opp = new Entity("opportunity");
                    opp["name"] = /*"queuing-" +*/ product["name"].ToString();
                    opp["bsd_project"] = product["bsd_projectcode"];
                    EntityReference pricelist_id = null;
                    if (product.Attributes.Contains("bsd_phaseslaunchid"))
                    {
                        opp["bsd_phaselaunch"] = product["bsd_phaseslaunchid"];
                        EntityReference priceList = PhasesLaunchPriceList((EntityReference)product["bsd_phaseslaunchid"]);
                        pricelist_id = priceList;
                        if (priceList != null)
                            opp["pricelevelid"] = priceList;
                        else
                            throw new Exception("Please choose pricelist for this phaseslaunch!");
                        opp["bsd_queuingfee"] = new Money(0);
                    }
                    else
                    {
                        EntityReference proRef = (EntityReference)product["bsd_projectcode"];
                        Entity project = service.Retrieve(proRef.LogicalName, proRef.Id, new ColumnSet(new string[] { "bsd_pricelistdefault", "bsd_bookingfee" }));
                        if (project == null)
                            throw new Exception("Project named '" + proRef.Name + "' is not available!");
                        //if (project.Attributes.Contains("bsd_pricelistdefault"))
                        //    opp["pricelevelid"] = project["bsd_pricelistdefault"];
                        //else throw new Exception("Please select price list default on project named '" + proRef.Name + "'");
                        opp["bsd_queuingfee"] = product.Contains("bsd_queuingfee") ? product["bsd_queuingfee"] : project.Contains("bsd_bookingfee") ? project["bsd_bookingfee"] : new Money(0);
                    }
                    EntityReference pricelist_ref = null;
                
                    if (pricelist_id != null)
                    {
                        var rplCopy = getListByIDCopy(service, pricelist_id.Id);
                       
                        if (rplCopy == null || rplCopy.Entities.Count == 0)
                        {
                        }
                        else
                        {
                            var copy = rplCopy[0];
                            pricelist_ref = new EntityReference(copy.LogicalName, copy.Id);
                        }
                        opp["bsd_pricelistapply"] = pricelist_ref;
                    }
                   
                   

                    Guid oppGuid = service.Create(opp);


                    Entity oppp = new Entity("opportunityproduct");
                    oppp["opportunityid"] = oppp["bsd_booking"] = new EntityReference("opportunity", oppGuid);
                    oppp["uomid"] = product["defaultuomid"];
                    oppp["bsd_floor"] = product["bsd_floor"];
                    oppp["bsd_block"] = product["bsd_blocknumber"];

                    oppp["bsd_project"] = product["bsd_projectcode"];
                    oppp["productid"] = oppp["bsd_units"] = productRef;
                    oppp["isproductoverridden"] = false;
                    oppp["ispriceoverridden"] = false;

                    if (opp.Contains("pricelevelid"))
                        oppp["bsd_pricelist"] = opp["pricelevelid"];

                    oppp["quantity"] = (decimal)1;
                    if (pricelist_ref != null)
                    {
                        Entity pricecopy = service.Retrieve(pricelist_ref.LogicalName, pricelist_ref.Id, new ColumnSet(true));
                        oppp["priceperunit"] = pricecopy["bsd_listprice"];
                    }
                    else 
                    {
                       if(product.Attributes.Contains("bsd_listprice"))
                          oppp["priceperunit"] = product["bsd_listprice"];
                    }
                       
                    if (product.Attributes.Contains("bsd_phaseslaunchid"))
                    {
                        oppp["bsd_status"] = true;
                        oppp["bsd_phaseslaunch"] = product["bsd_phaseslaunchid"];
                    }

                    service.Create(oppp);
                    content = "tmp={type:'Success',content:'" + oppGuid.ToString() + "'}";
                }
                context.OutputParameters["Result"] = content;
                #endregion
            }
            else if (command == "Reservation")
            {
                throw new Exception("hello");
                #region Reservation

                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);
                //Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet(
                //new string[] { "name", "statuscode", "defaultuomid", "bsd_listprice", "statecode", "bsd_projectcode", "bsd_blocknumber", "bsd_floor", "bsd_phaseslaunchid" }));

                Entity product = RetrieveValidUnit(productRef.Id);
                if (product == null)
                    throw new Exception("Unit is not avaliable please check detail of unit!");
                if (((OptionSetValue)product["statecode"]).Value == 1)
                    throw new Exception("This unit is not public!");
                else if (((OptionSetValue)product["statuscode"]).Value != 100000000 && ((OptionSetValue)product["statuscode"]).Value != 100000001 && ((OptionSetValue)product["statuscode"]).Value != 100000004)
                    throw new Exception("Unit must be available or on hold!");
                else if (((OptionSetValue)product["statuscode"]).Value == 100000002)
                    throw new Exception("Unit is sold!");

                if (!product.Attributes.Contains("bsd_phaseslaunchid"))
                    throw new Exception("Unit is not launched!");
                if (!product.Attributes.Contains("bsd_floor"))
                    throw new Exception("Please select floor for this unit!");
                if (!product.Attributes.Contains("bsd_blocknumber"))
                    throw new Exception("Please select block for this unit!");
                if (!product.Attributes.Contains("bsd_projectcode"))
                    throw new Exception("Please select project for this unit!");
                if (!product.Attributes.Contains("defaultuomid"))
                    throw new Exception("Please select default unit for this unit!");
                if (!product.Attributes.Contains("bsd_depositamount"))
                    throw new Exception("Please provide deposit for this unit!");


                Entity quote = new Entity("quote");
                quote["name"] = product["name"];
                quote["bsd_projectid"] = product["bsd_projectcode"];
                quote["transactioncurrencyid"] = product["transactioncurrencyid"];
                //if (product.Attributes.Contains("pj.bsd_defaultpaymentscheme"))
                //    quote["bsd_paymentscheme"] = ((AliasedValue)product["pj.bsd_defaultpaymentscheme"]).Value;
                quote["bsd_depositfee"] = product["bsd_depositamount"];
                quote["bsd_phaseslaunchid"] = product["bsd_phaseslaunchid"];
                if (product.Attributes.Contains("bsd_phaseslaunchid"))
                {
                    EntityReference priceList = PhasesLaunchPriceList((EntityReference)product["bsd_phaseslaunchid"]);
                    if (priceList != null)
                        quote["bsd_pricelistphaselaunch"] = priceList;
                    else
                        throw new Exception("Please choose pricelist for this phaseslaunch!");

                }
                else if (product.Attributes.Contains("pricelevelid"))
                    quote["bsd_pricelistphaselaunch"] = product["pricelevelid"];
                else
                    throw new Exception("Please enter 'default price list' on this Unit!");
                #region Update pricelist mới nhất
                EntityReference pricelist_ref = null;
                pricelist_ref = (EntityReference)quote["pricelevelid"];
                if (pricelist_ref != null)
                {
                    var rplCopy = getListByIDCopy(service, pricelist_ref.Id);
                    if (rplCopy == null || rplCopy.Entities.Count == 0)
                    {
                    }
                    else
                    {
                        var copy = rplCopy[0];
                        pricelist_ref = new EntityReference(copy.LogicalName, copy.Id);
                    }
                    quote["pricelevelid"] = pricelist_ref;
                }
                #endregion

                //get customer fomr queue
                if (context.InputParameters.Contains("Parameters") && context.InputParameters["Parameters"] != null)
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(InputParameter[]));
                    MemoryStream str = new MemoryStream(Encoding.UTF8.GetBytes((string)context.InputParameters["Parameters"]));
                    InputParameter[] parameters = (InputParameter[])ser.ReadObject(str);
                    foreach (InputParameter i in parameters)
                    {
                        if (i.action == command)
                        {
                            //throw new Exception("hello");
                            Entity opp = service.Retrieve(i.name, Guid.Parse(i.value), new ColumnSet(new string[] {
                                "customerid"
                            }));
                            EntityReference cusRef = opp.Contains("customerid") ? (EntityReference)opp["customerid"] : null;
                            if (cusRef != null)
                                quote["customerid"] = cusRef;
                        }
                    }
                }
                Guid quoteId = service.Create(quote);
                Entity qtmp = service.Retrieve(quote.LogicalName, quoteId, new ColumnSet(new string[] { "createdon" }));
                DateTime t = (DateTime)qtmp["createdon"];
                qtmp.Attributes.Clear();
                qtmp["bsd_reservationtime"] = t;
                service.Update(qtmp);

                Entity quoteProduct = new Entity("quotedetail");
                quoteProduct["isproductoverridden"] = true;
                quoteProduct["ispriceoverridden"] = true;
                quoteProduct["productid"] = new EntityReference("product", productRef.Id);
                quoteProduct["quantity"] = (decimal)2;
                quoteProduct["priceperunit"] = new Money (((Money) product["bsd_price"]).Value); 
                quoteProduct["uomid"] = product["defaultuomid"];
                quoteProduct["transactioncurrencyid"] = product["transactioncurrencyid"];
                quoteProduct["quoteid"] = new EntityReference("quote", quoteId);
                service.Create(quoteProduct);
                throw new Exception("hello" + ((decimal) quoteProduct["quantity"]));
                context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + quoteId.ToString() + "'}";
                #endregion
            }
            else if (command == "OptionEntry")
            {
                #region OptionEntry

                factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = factory.CreateOrganizationService(context.UserId);

                Entity product = RetrieveValidUnit(productRef.Id);
                if (product == null)
                    throw new Exception("Unit is not avaliable please check detail of unit!");
                if (((OptionSetValue)product["statecode"]).Value == 1)
                    throw new Exception("This unit is not public!");
                else if (((OptionSetValue)product["statuscode"]).Value != 100000000 && ((OptionSetValue)product["statuscode"]).Value != 100000001)
                    throw new Exception("Unit must be available or on hold!");
                else if (!product.Attributes.Contains("bsd_phaseslaunchid"))
                    throw new Exception("Unit is not launched");
                else if (!product.Attributes.Contains("bsd_floor"))
                    throw new Exception("Please select floor for this unit!");
                else if (!product.Attributes.Contains("bsd_blocknumber"))
                    throw new Exception("Please select block for this unit!");
                else if (!product.Attributes.Contains("bsd_projectcode"))
                    throw new Exception("Please select project for this unit!");
                else if (!product.Attributes.Contains("defaultuomid"))
                    throw new Exception("Please select default unit for this unit!");

                Entity order = new Entity("salesorder");
                order["name"] = product["productnumber"];
                order["transactioncurrencyid"] = product["transactioncurrencyid"];
                order["bsd_project"] = product["bsd_projectcode"];
                order["bsd_phaseslaunch"] = product["bsd_phaseslaunchid"];
                order["statuscode"] = new OptionSetValue(100000008);
                if (product.Attributes.Contains("bsd_phaseslaunchid"))
                {
                    EntityReference priceList = PhasesLaunchPriceList((EntityReference)product["bsd_phaseslaunchid"]);
                    if (priceList != null)
                        order["pricelevelid"] = priceList;
                    else
                        throw new Exception("Please choose pricelist for this phaseslaunch!");
                }
                else if (product.Attributes.Contains("pricelevelid"))
                    order["pricelevelid"] = product["pricelevelid"];
                else
                    throw new Exception("Please enter 'default price list' on this Unit!");
                Guid orderId = service.Create(order);

                Entity orderDetail = new Entity("salesorderdetail");
                orderDetail.Id = Guid.NewGuid();
                orderDetail["isproductoverridden"] = false;
                orderDetail["ispriceoverridden"] = false;
                orderDetail["productid"] = new EntityReference("product", productRef.Id);
                orderDetail["quantity"] = (decimal)1;
                orderDetail["uomid"] = product["defaultuomid"];
                orderDetail["transactioncurrencyid"] = product["transactioncurrencyid"];
                orderDetail["salesorderid"] = new EntityReference("salesorder", orderId);
                service.Create(orderDetail);
                context.OutputParameters["Result"] = "tmp={type:'Success',content:'" + orderId.ToString() + "'}";
                #endregion
            }
            else
                throw new Exception("Command is not valid!");
        }

        public EntityCollection getListByIDCopy(IOrganizationService service, Guid idcopy)
        {
            #region --- Danh sách sắp xếp theo filter createdon mới nhất top 1 ---
            var xml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='pricelevel'>
                    <attribute name='name' />
                    <attribute name='transactioncurrencyid' />
                    <attribute name='enddate' />
                    <attribute name='begindate' />
                    <attribute name='statecode' />
                    <attribute name='createdon' />
                    <attribute name='pricelevelid' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                      <condition attribute='bsd_approved' operator='eq' value='1' />
                      <condition attribute='bsd_pricelistcopy' operator='eq'  uitype='pricelevel' value='" + idcopy + @"' />
                    </filter>
                  </entity>
                </fetch>";
            #endregion

            var rs = service.RetrieveMultiple(new QueryExpression(xml.ToString()));
            return rs;
        }

        private Entity RetrieveValidUnit(Guid unitId)
        {
            QueryExpression q = new QueryExpression();
            q.EntityName = "product";
            q.ColumnSet = new ColumnSet(new string[] {
            "name","productnumber","productid","bsd_floor","bsd_blocknumber","bsd_queuingfee"
            ,"bsd_projectcode","bsd_phaseslaunchid","defaultuomid"
            ,"statuscode","statecode","transactioncurrencyid","pricelevelid","bsd_depositamount"
            });
            q.Criteria = new FilterExpression(LogicalOperator.And);
            q.Criteria.AddCondition(new ConditionExpression("productid", ConditionOperator.Equal, unitId));
            LinkEntity link_floor_unit = new LinkEntity("product", "bsd_floor", "bsd_floor", "bsd_floorid", JoinOperator.Inner);
            link_floor_unit.EntityAlias = "fl";
            link_floor_unit.Columns = new ColumnSet(new string[] { "bsd_block" });
            q.LinkEntities.Add(link_floor_unit);
            LinkEntity link_block_floor = new LinkEntity("bsd_floor", "bsd_block", "bsd_block", "bsd_blockid", JoinOperator.Inner);
            link_block_floor.EntityAlias = "bl";
            link_block_floor.Columns = new ColumnSet(new string[] { "bsd_project" });
            link_floor_unit.LinkEntities.Add(link_block_floor);
            LinkEntity link_project_block = new LinkEntity("bsd_block", "bsd_project", "bsd_project", "bsd_projectid", JoinOperator.Inner);
            link_project_block.EntityAlias = "pj";
            link_project_block.Columns = new ColumnSet(new string[] { "bsd_defaultpaymentscheme", "bsd_pricelistdefault" });
            link_block_floor.LinkEntities.Add(link_project_block);
            q.TopCount = 1;

            #region FetchXML

            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("");
            //sb.AppendLine("<fetch mapping='logical' count='1' output-format='xml-platform'>");
            //sb.AppendLine("<entity name='product'>");
            //sb.AppendLine("<attribute name='name'/>");
            //sb.AppendLine("<attribute name='productid'/>");
            //sb.AppendLine("<attribute name='bsd_floor'/>");
            //sb.AppendLine("<attribute name='bsd_blocknumber'/>");
            //sb.AppendLine("<attribute name='bsd_projectcode'/>");
            //sb.AppendLine("<attribute name='bsd_phaseslaunchid'/>");
            //sb.AppendLine("<attribute name='bsd_listprice'/>");
            //sb.AppendLine("<attribute name='defaultuomid'/>");
            //sb.AppendLine("<attribute name='statuscode'/>");
            //sb.AppendLine("<attribute name='statecode'/>");
            //sb.AppendLine("<filter type='and'>");
            //sb.AppendLine("<condition attribute='productid' operator='eq' value='" + unitId.ToString() + "'></condition>");
            //sb.AppendLine("</filter>");
            //sb.AppendLine("<link-entity name='bsd_floor' from='bsd_floorid' to='bsd_floor' link-type='inner'>");
            //sb.AppendLine("<attribute name='bsd_block'/>");
            //sb.AppendLine("<link-entity name='bsd_block' from='bsd_blockid' to='bsd_block' link-type='inner'>");
            //sb.AppendLine("<attribute name='bsd_project'/>");
            //sb.AppendLine("<link-entity name='bsd_project' from='bsd_projectid' to='bsd_project' link-type='inner'></link-entity>");
            //sb.AppendLine("</link-entity>");
            //sb.AppendLine("</link-entity>");
            //sb.AppendLine("</entity>");
            //sb.AppendLine("</fetch>"); 
            #endregion
            //EntityCollection entcs = service.RetrieveMultiple(new FetchExpression(sb.ToString()));
            EntityCollection entcs = service.RetrieveMultiple(q);
            if (entcs.Entities.Count == 0)
                return null;
            else
                return entcs.Entities[0];
        }
        //queuefee,depositfee
        private EntityReference PhasesLaunchPriceList(EntityReference phaseLaunch)
        {
            Entity en = service.Retrieve(phaseLaunch.LogicalName, phaseLaunch.Id, new ColumnSet(new string[] { "bsd_pricelistid" }));
            if (en.Attributes.Contains("bsd_pricelistid"))
                return (EntityReference)en["bsd_pricelistid"];
            else
                return null;
        }

        private Money GetQueuefee(EntityReference phaseF)
        {
            Money m = new Money(0);
            if (phaseF != null)
            {
                Entity tmp = service.Retrieve(phaseF.LogicalName, phaseF.Id, new ColumnSet(new string[] {
                    "bsd_bookingfee"
                }));
                if (tmp.Contains("bsd_bookingfee"))
                    m = (Money)tmp["bsd_bookingfee"];
            }
            return m;
        }

        private Money GetQepositfee(EntityReference pmSchRef)
        {
            Money m = new Money(0);
            if (pmSchRef != null)
            {
                Entity tmp = service.Retrieve(pmSchRef.LogicalName, pmSchRef.Id, new ColumnSet(new string[] {
                    "bsd_depositamount"
                }));
                if (tmp.Contains("bsd_depositamount"))
                    m = (Money)tmp["bsd_depositamount"];
            }
            return m;
        }

        [DataContract]
        public class InputParameter
        {
            [DataMember]
            public string action { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string value { get; set; }
        }
    }
}
