using Plugin_QueryBuilderGroup_Create_Update.Models;
using System;

namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public class FieldPathResolver
    {
        public FieldPath Resolve(string field)
        {
            var result = new FieldPath();

            // no relationship
            if (!field.Contains("|"))
            {
                var parts = field.Split('.');
                result.HasRelationship = false;
                result.TargetAttribute = parts[1];
                return result;
            }

            //----------------------------------
            // FORMAT YOU ARE USING:
            //
            // root.targetEntity |
            // lookup.attribute
            //----------------------------------

            var split = field.Split('|');

            var left = split[0].Split('.');
            var right = split[1].Split('.');

            if (left.Length != 2 || right.Length != 2)
                throw new Exception("Invalid field format");

            result.HasRelationship = true;

            // entity cần join
            result.TargetEntity = left[1];

            // lookup nằm trên root
            result.LookupField = right[0];

            // attribute filter
            result.TargetAttribute = right[1];

            return result;
        }
    }
}
