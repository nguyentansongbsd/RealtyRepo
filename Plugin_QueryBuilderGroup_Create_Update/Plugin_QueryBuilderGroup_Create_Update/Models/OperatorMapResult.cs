namespace Plugin_QueryBuilderGroup_Create_Update.Models
{
    public class OperatorMapResult
    {
        public string FetchOperator;
        public string Prefix = "";
        public string Suffix = "";
        public bool RequireValue = true;
        public bool IsBetween { get; set; }
        public bool IsNotBetween { get; set; }
    }
}
