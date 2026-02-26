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
              <attribute name='{rootEntity}id' />
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
            foreach (var node in group.Rules)
            {
                if (node is QueryGroup g)
                {
                    filter.Append($"<filter type='{g.Condition}'>");

                    CompileGroup(g, filter, manager);

                    filter.Append("</filter>");
                }

                if (node is QueryRule r)
                {
                    BuildCondition(r, filter, manager);
                }
            }
        }

        // =====================================================
        // CONDITION BUILDER
        // =====================================================
        private void BuildCondition(
            QueryRule rule,
            StringBuilder filter,
            LinkEntityManager manager)
        {
            var path = new FieldPathResolver()
                .Resolve(rule.Field);

            var map = OperatorMapper.Map(rule.Operator);

            string conditionXml = "";

            // =====================================================
            // BETWEEN / NOT BETWEEN
            // =====================================================
            if (map.IsBetween || map.IsNotBetween)
            {
                if (!(rule.Value is JArray arr) || arr.Count != 2)
                    throw new System.Exception(
                        "Between operator requires 2 values");

                string v1 = arr[0].ToString();
                string v2 = arr[1].ToString();

                conditionXml =
                    map.IsBetween
                    ? $@"
                <filter type='and'>
                   <condition attribute='{path.TargetAttribute}' operator='ge' value='{v1}'/>
                   <condition attribute='{path.TargetAttribute}' operator='le' value='{v2}'/>
                </filter>"
                                    : $@"
                <filter type='or'>
                   <condition attribute='{path.TargetAttribute}' operator='lt' value='{v1}'/>
                   <condition attribute='{path.TargetAttribute}' operator='gt' value='{v2}'/>
                </filter>";
            }

            // =====================================================
            // IN / NOT IN  ✅ CRM FORMAT
            // =====================================================
            else if (map.FetchOperator == "in"
                  || map.FetchOperator == "not-in")
            {
                if (!(rule.Value is JArray arr))
                    throw new System.Exception(
                        "IN operator requires array value");

                var sb = new StringBuilder();

                sb.Append(
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}'>");

                foreach (var v in arr)
                    sb.Append($"<value>{v}</value>");

                sb.Append("</condition>");

                conditionXml = sb.ToString();
            }

            // =====================================================
            // NULL / NOT NULL
            // =====================================================
            else if (!map.RequireValue)
            {
                conditionXml =
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}'/>";
            }

            // =====================================================
            // LIKE PREFIX SUFFIX
            // =====================================================
            else
            {
                string value = rule.Value?.ToString() ?? "";

                if (!string.IsNullOrEmpty(map.Prefix))
                    value = map.Prefix + value;

                if (!string.IsNullOrEmpty(map.Suffix))
                    value += map.Suffix;

                conditionXml =
                    $"<condition attribute='{path.TargetAttribute}' operator='{map.FetchOperator}' value='{value}'/>";
            }

            // =====================================================
            // ROOT vs LINK ENTITY
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