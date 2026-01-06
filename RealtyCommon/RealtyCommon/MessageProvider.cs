using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RealtyCommon
{
    public static class MessageProvider
    {
        public static string GetMessage(IOrganizationService service, IPluginExecutionContext context, string notificationName, IDictionary<string, object> parameters = null)
        {
            int userLanguageCode = GetLanguageCode(service, context);

            var fetchXml = $@"<?xml version=""1.0"" encoding=""utf-16""?>
            <fetch>
              <entity name=""bsd_messageconfig"">
                <attribute name=""bsd_name"" />
                <filter>
                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                  <condition attribute=""bsd_name"" operator=""not-null"" />
                </filter>
                <link-entity name=""bsd_notificationconfig"" from=""bsd_notificationconfigid"" to=""bsd_notificationconfig"">
                  <filter>
                    <condition attribute=""bsd_name"" operator=""eq"" value=""{notificationName}"" />
                  </filter>
                </link-entity>
                <link-entity name=""bsd_languagecode"" from=""bsd_languagecodeid"" to=""bsd_language"">
                  <filter>
                    <condition attribute=""bsd_code"" operator=""eq"" value=""{userLanguageCode}"" />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";
            EntityCollection rs = service.RetrieveMultiple(new FetchExpression(fetchXml));
            if (rs != null && rs.Entities != null && rs.Entities.Count > 0)
            {
                string msg = (string)rs.Entities[0]["bsd_name"];

                if (parameters != null && parameters.Count > 0)
                {
                    foreach (var kv in parameters)
                    {
                        msg = msg.Replace("{" + kv.Key + "}", kv.Value?.ToString());
                    }
                }

                return msg;
            }

            return "Message not found, please check 'Message Config' again.";
        }

        private static int GetLanguageCode(IOrganizationService service, IPluginExecutionContext context)
        {
            Entity userSettings = service.RetrieveMultiple(
                                    new QueryExpression("usersettings")
                                    {
                                        ColumnSet = new ColumnSet("uilanguageid"),
                                        Criteria = new FilterExpression
                                        {
                                            Conditions = {
                                                new ConditionExpression("systemuserid", ConditionOperator.Equal, context.UserId)
                                            }
                                        }
                                    }).Entities.FirstOrDefault();

            // language code: https://crmminds.com/2023/03/08/dynamics-365-language-codes/
            int userLanguageId = 1033; // default EN
            if (userSettings != null && userSettings.Contains("uilanguageid"))
            {
                userLanguageId = (int)userSettings["uilanguageid"];
            }
            return userLanguageId;
        }
    }
}
