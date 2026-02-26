using System.Collections.Generic;


namespace Plugin_QueryBuilderGroup_Create_Update.Models
{
    public class QueryGroup : QueryNode
    {
        public string Condition { get; set; }
        public List<QueryNode> Rules { get; set; }
    }
}
