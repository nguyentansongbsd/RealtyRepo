using System.Collections.Generic;
using System.Text;

namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public class LinkEntityManager
    {
        private class LinkNode
        {
            public string Entity;
            public string Lookup;
            public string Alias;

            public List<string> Conditions =
                new List<string>();
        }

        private readonly Dictionary<string, LinkNode> _links
            = new Dictionary<string, LinkNode>();

        private int aliasIndex = 1;

        public void AddCondition(
            string entity,
            string lookup,
            string condition)
        {
            string key = entity + "|" + lookup;

            if (!_links.ContainsKey(key))
            {
                _links[key] = new LinkNode
                {
                    Entity = entity,
                    Lookup = lookup,
                    Alias = "l" + aliasIndex++
                };
            }

            _links[key].Conditions.Add(condition);
        }

        public string Build()
        {
            var sb = new StringBuilder();

            foreach (var link in _links.Values)
            {
                sb.Append(
$@"<link-entity name='{link.Entity}'
    from='{link.Entity}id'
    to='{link.Lookup}'
    alias='{link.Alias}'
    link-type='inner'>");

                sb.Append("<filter type='and'>");

                foreach (var cond in link.Conditions)
                {
                    // tránh filter lồng filter
                    sb.Append(cond
                        .Replace("<filter type='and'>", "")
                        .Replace("</filter>", ""));
                }

                sb.Append("</filter>");
                sb.Append("</link-entity>");
            }

            return sb.ToString();
        }
    }
}