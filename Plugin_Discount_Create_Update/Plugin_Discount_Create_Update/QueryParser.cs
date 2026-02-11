using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Plugin_Discount_Create_Update
{
    public static class QueryParser
    {
        public static string Convert(string input, string rootEntity)
        {
            rootEntity = StringHelper.Clean(rootEntity);

            var ast = Parse(input, rootEntity);

            var entity = new XElement("entity",
                new XAttribute("name", rootEntity));

            var filter = BuildFilter(ast, entity);
            if (filter != null)
                entity.Add(filter);

            var fetch = new XElement("fetch", entity);
            return fetch.ToString();
        }

        // ================= BUILD FETCH =================

        private static XElement BuildFilter(QueryNode node, XElement root)
        {
            if (node is LogicalNode ln)
            {
                var filter = new XElement("filter",
                    new XAttribute("type", ln.Operator));

                foreach (var c in ln.Children)
                {
                    var child = BuildFilter(c, root);
                    if (child != null)
                        filter.Add(child);
                }

                return filter;
            }

            if (node is ConditionNode cn)
            {
                // ROOT CONDITION
                if (cn.Link == null)
                {
                    var cond = new XElement("condition",
                        new XAttribute("attribute", StringHelper.Clean(cn.Attribute)),
                        new XAttribute("operator", cn.Operator));

                    foreach (var v in cn.Values)
                        cond.Add(new XElement("value", StringHelper.Clean(v)));

                    return cond;
                }

                // LINK-ENTITY CONDITION
                var le = new XElement("link-entity",
                    new XAttribute("name", StringHelper.Clean(cn.Link.EntityName)),
                    new XAttribute("from", StringHelper.Clean(cn.Link.From)),
                    new XAttribute("to", StringHelper.Clean(cn.Link.To)),
                    new XAttribute("alias", cn.Link.Alias),
                    new XAttribute("link-type", cn.Link.JoinType)
                );

                var filter = new XElement("filter",
                    new XAttribute("type", "and"));

                var condInLink = new XElement("condition",
                    new XAttribute("attribute", StringHelper.Clean(cn.Attribute)),
                    new XAttribute("operator", cn.Operator));

                foreach (var v in cn.Values)
                    condInLink.Add(new XElement("value", StringHelper.Clean(v)));

                filter.Add(condInLink);
                le.Add(filter);

                root.Add(le);

                return null; // QUAN TRỌNG: link-entity đã add vào root
            }

            throw new Exception("Invalid AST");
        }

        // ================= PARSE =================

        public static QueryNode Parse(string input, string rootEntity)
        {
            if (input.Contains(" AND "))
            {
                var parts = input.Split(new[] { " AND " }, StringSplitOptions.None);
                var node = new LogicalNode { Operator = "and" };
                foreach (var p in parts)
                    node.Children.Add(Parse(p.Trim(), rootEntity));
                return node;
            }

            if (input.Contains(" OR "))
            {
                var parts = input.Split(new[] { " OR " }, StringSplitOptions.None);
                var node = new LogicalNode { Operator = "or" };
                foreach (var p in parts)
                    node.Children.Add(Parse(p.Trim(), rootEntity));
                return node;
            }

            return ParseCondition(input, rootEntity);
        }

        private static ConditionNode ParseCondition(string input, string rootEntity)
        {
            var parts = input.Split(new[] { " IN " }, StringSplitOptions.None);
            if (parts.Length != 2)
                throw new Exception("Invalid condition: " + input);

            var left = StringHelper.Clean(parts[0]);
            var valuesRaw = parts[1].Trim().Trim('(', ')');

            var values = valuesRaw
                .Split(',')
                .Select(v => StringHelper.Clean(v.Trim().Trim('\'')))
                .ToList();

            var node = new ConditionNode
            {
                RootEntity = StringHelper.Clean(rootEntity),
                Operator = "in",
                Values = values
            };

            // remove root entity prefix
            if (left.StartsWith(rootEntity + "."))
                left = left.Substring(rootEntity.Length + 1);

            left = StringHelper.Clean(left);

            // CASE 1: ROOT FIELD
            if (!left.Contains("|"))
            {
                node.Attribute = left;
                return node;
            }

            // CASE 2: LINK-ENTITY 1 cấp
            // bsd_product|bsd_productid.bsd_blocknumber
            var linkParts = left.Split('|');

            var entityName = StringHelper.Clean(linkParts[0]);

            var restParts = linkParts[1].Split('.');
            var lookupField = StringHelper.Clean(restParts[0]);
            var attribute = StringHelper.Clean(restParts[1]);

            node.Attribute = attribute;

            node.Link = new LinkEntityInfo
            {
                EntityName = entityName,
                From = lookupField,
                To = lookupField,
                Alias = "U",
                JoinType = "inner"
            };

            return node;
        }

        // ================= EXTRACT ROOT =================

        public static string ExtractRootEntity(string input)
        {
            var first = StringHelper.Clean(input);

            if (first.Contains(" AND "))
                first = first.Split(new[] { " AND " }, StringSplitOptions.None)[0];
            if (first.Contains(" OR "))
                first = first.Split(new[] { " OR " }, StringSplitOptions.None)[0];

            var idx = first.IndexOf(" IN ", StringComparison.OrdinalIgnoreCase);
            if (idx > 0)
                first = first.Substring(0, idx);

            return StringHelper.Clean(first.Split('.')[0]);
        }
    }
}
