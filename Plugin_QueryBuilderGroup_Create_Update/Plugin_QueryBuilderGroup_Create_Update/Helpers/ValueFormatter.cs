using System;


namespace Plugin_QueryBuilderGroup_Create_Update.Helpers
{
    public static class ValueFormatter
    {
        public static string Format(object val,
                                    string fieldType)
        {
            if (fieldType == "DateTime")
            {
                DateTime dt =
                  DateTime.Parse(val.ToString());

                return dt.ToUniversalTime()
                         .ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            return val.ToString();
        }
    }
}
