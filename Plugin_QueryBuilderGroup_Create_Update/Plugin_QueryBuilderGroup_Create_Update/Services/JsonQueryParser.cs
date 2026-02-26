using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Plugin_QueryBuilderGroup_Create_Update.Models;
namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public class JsonQueryParser
    {
        public QueryGroup Parse(string json)
        {
            JObject obj = JObject.Parse(json);
            return ParseGroup(obj);
        }

        private QueryGroup ParseGroup(JObject obj)
        {
            var group = new QueryGroup
            {
                Condition = obj["condition"].ToString(),
                Rules = new List<QueryNode>()
            };

            foreach (var r in obj["rules"])
            {
                if (r["rules"] != null)
                    group.Rules.Add(ParseGroup((JObject)r));
                else
                    group.Rules.Add(r.ToObject<QueryRule>());
            }

            return group;
        }
    }
}
