namespace Plugin_QueryBuilderGroup_Create_Update.Models
{
    public class FieldPath
    {
        public string SourceEntity;
        public string LookupField;

        public string TargetEntity;
        public string TargetAttribute;

        public bool HasRelationship;
    }
}
