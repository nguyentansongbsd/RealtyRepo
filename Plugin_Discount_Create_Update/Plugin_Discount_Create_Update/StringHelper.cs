namespace Plugin_Discount_Create_Update
{
    internal static class StringHelper
    {
        internal static string Clean(string input)
        {
            return input?.Trim().Trim('`');
        }
    }
}
