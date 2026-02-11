using System.Collections.Generic;

namespace Plugin_Discount_Create_Update
{
    public class LogicalNode : QueryNode
    {
        public string Operator; // "and" | "or"
        public List<QueryNode> Children = new List<QueryNode>();
    }
}
