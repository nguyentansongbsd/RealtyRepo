namespace Plugin_QueryBuilderGroup_Create_Update.Models
{
    public class QueryRule : QueryNode
    {
        public string Field { get; set; }
        public string Operator { get; set; }
        public object Value { get; set; }
        public string FieldType { get; set; }
    }
}
