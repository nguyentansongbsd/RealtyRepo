using System;
using System.Text;
using Newtonsoft.Json.Linq;
using Plugin_QueryBuilderGroup_Create_Update.Models;

namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public class FetchXmlCompiler
    {
        public string Build(
            QueryGroup group,
            string rootEntity)
        {
            var entityFilter = new StringBuilder();
            var linkManager = new LinkEntityManager();

            CompileGroup(group, entityFilter, linkManager);

            return $@"
<fetch>
  <entity name='{rootEntity}'>
    <attribute name='{rootEntity}id'/>
    <filter type='{group.Condition}'>
        {entityFilter}
    </filter>
    {linkManager.Build()}
  </entity>
</fetch>";
        }

        // =====================================================
        // GROUP
        // =====================================================
        private void CompileGroup(
    QueryGroup group,
    StringBuilder filter,
    LinkEntityManager manager)
        {
            var inner = new StringBuilder();

            foreach (var node in group.Rules)
            {
                if (node is QueryGroup g)
                {
                    CompileGroup(g, inner, manager);
                }
                else if (node is QueryRule r)
                {
                    BuildCondition(r, inner, manager);
                }
            }

            if (inner.Length > 0)
            {
                filter.Append(
                    $"<filter type='{group.Condition.ToLower()}'>");

                filter.Append(inner);

                filter.Append("</filter>");
            }
        }

        // =====================================================
        // DATE FORMAT FIX (🔥 QUAN TRỌNG)
        // =====================================================
        private string FormatDate(object value, bool endOfDay = false)
        {
            if (value == null) return "";

            DateTime dt = Convert.ToDateTime(value);

            //if (endOfDay)
                dt = dt.Date.AddDays(1).AddSeconds(-1);
            //else
            //    dt = dt.Date;

            return dt.ToUniversalTime()
                     .ToString("yyyy-MM-ddTHH:mm:ssZ");
        }

        // =====================================================
        // CONDITION BUILDER
        // =====================================================
        private void BuildCondition(
            QueryRule rule,
            StringBuilder filter,
            LinkEntityManager manager)
        {
            var path =
                new FieldPathResolver()
                .Resolve(rule.Field);

            var map =
                OperatorMapper.Map(rule.Operator);

            string conditionXml = "";

            bool isDate =
                rule.FieldType?.ToLower() == "date";

            // =====================================================
            // BETWEEN / NOT BETWEEN
            // =====================================================
            if (map.IsBetween || map.IsNotBetween)
            {
                if (!(rule.Value is JArray arr) || arr.Count != 2)
                    throw new Exception("Between requires 2 values");

                string v1 = isDate
                    ? FormatDate(arr[0].ToString(), false)
                    : arr[0].ToString();

                string v2 = isDate
                    ? FormatDate(arr[1].ToString(), true)
                    : arr[1].ToString();

                if (map.IsBetween)
                {
                    conditionXml +=
                        $"<condition attribute='{path.TargetAttribute}' operator='ge' value='{v1}'/>";
                    conditionXml +=
                        $"<condition attribute='{path.TargetAttribute}' operator='le' value='{v2}'/>";
                }
                else
                {
                    conditionXml +=
                        $"<condition attribute='{path.TargetAttribute}' operator='lt' value='{v1}'/>";
                    conditionXml +=
                        $"<condition attribute='{path.TargetAttribute}' operator='gt' value='{v2}'/>";
                }
            }

            // =====================================================
            // IN / NOT IN
            // =====================================================
            else if (map.FetchOperator == "in"
                  || map.FetchOperator == "not-in")
            {
                if (!(rule.Value is JArray arr))
                    throw new Exception("IN requires array");

                var sb = new StringBuilder();

                sb.Append(
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}'>");

                foreach (var v in arr)
                    sb.Append($"<value>{v}</value>");

                sb.Append("</condition>");

                conditionXml = sb.ToString();
            }

            // =====================================================
            // NULL
            // =====================================================
            else if (!map.RequireValue)
            {
                conditionXml =
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}'/>";
            }

            // =====================================================
            // NORMAL CONDITION
            // =====================================================
            else
            {
                string value;

                if (isDate)
                {
                    // 🔥 FIX <= DATE MẤT RECORD
                    if (map.FetchOperator == "le")
                        value = FormatDate(rule.Value, true);
                    else
                        value = FormatDate(rule.Value);
                }
                else
                {
                    value = rule.Value?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(map.Prefix))
                        value = map.Prefix + value;

                    if (!string.IsNullOrEmpty(map.Suffix))
                        value += map.Suffix;
                }

                conditionXml =
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}' value='{value}'/>";
            }

            // =====================================================
            // ROOT OR LINK
            // =====================================================
            if (!path.HasRelationship)
            {
                filter.Append(conditionXml);
            }
            else
            {
                manager.AddCondition(
                    path.TargetEntity,
                    path.LookupField,
                    conditionXml);
            }
        }
    }
}