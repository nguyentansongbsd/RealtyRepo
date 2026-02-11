using System.Collections.Generic;

namespace Plugin_Discount_Create_Update
{
    public class ConditionNode : QueryNode
    {
        public string RootEntity;
        public string Attribute;
        public string Operator;
        public List<string> Values = new List<string>();
        public LinkEntityInfo Link; // chỉ 1 link-entity
    }
}
