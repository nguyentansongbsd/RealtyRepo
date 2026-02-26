using Plugin_QueryBuilderGroup_Create_Update.Models;
using System;

namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public static class OperatorMapper
    {
        public static OperatorMapResult Map(string op)
        {
            op = op.ToLower();

            switch (op)
            {
                case "equal":
                    return new OperatorMapResult { FetchOperator = "eq" };

                case "notequal":
                    return new OperatorMapResult { FetchOperator = "ne" };

                case "greaterthan":
                    return new OperatorMapResult { FetchOperator = "gt" };

                case "greaterthanorequal":
                    return new OperatorMapResult { FetchOperator = "ge" };

                case "lessthan":
                    return new OperatorMapResult { FetchOperator = "lt" };

                case "lessthanorequal":
                    return new OperatorMapResult { FetchOperator = "le" };

                case "contains":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Prefix = "%",
                        Suffix = "%"
                    };

                case "beginswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Suffix = "%"
                    };

                case "endswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Prefix = "%"
                    };

                case "in":
                    return new OperatorMapResult
                    {
                        FetchOperator = "in"
                    };

                case "notin":
                    return new OperatorMapResult
                    {
                        FetchOperator = "not-in"
                    };

                case "isnull":
                    return new OperatorMapResult
                    {
                        FetchOperator = "null",
                        RequireValue = false
                    };

                case "isnotnull":
                    return new OperatorMapResult
                    {
                        FetchOperator = "not-null",
                        RequireValue = false
                    };

                // ✅ BETWEEN
                case "between":
                    return new OperatorMapResult
                    {
                        IsBetween = true
                    };

                // ✅ NOT BETWEEN
                case "not-between":
                case "notbetween":
                    return new OperatorMapResult
                    {
                        IsNotBetween = true
                    };

                default:
                    throw new Exception("Operator not supported: " + op);
            }
        }
    }
}