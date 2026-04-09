using Plugin_QueryBuilderGroup_Create_Update.Models;
using System;

namespace Plugin_QueryBuilderGroup_Create_Update.Services
{
    public static class OperatorMapper
    {
        public static OperatorMapResult Map(string op)
        {
            if (string.IsNullOrWhiteSpace(op))
                throw new Exception("Operator is null or empty.");

            op = op.Trim().ToLower();

            switch (op)
            {
                case "equal":
                case "eq":
                    return new OperatorMapResult { FetchOperator = "eq" };

                case "notequal":
                case "ne":
                    return new OperatorMapResult { FetchOperator = "ne" };

                case "greaterthan":
                case "gt":
                    return new OperatorMapResult { FetchOperator = "gt" };

                case "greaterthanorequal":
                case "ge":
                    return new OperatorMapResult { FetchOperator = "ge" };

                case "lessthan":
                case "lt":
                    return new OperatorMapResult { FetchOperator = "lt" };

                case "lessthanorequal":
                case "le":
                    return new OperatorMapResult { FetchOperator = "le" };

                case "contains":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Prefix = "%",
                        Suffix = "%"
                    };

                case "doesnotcontain":
                case "notcontains":
                    return new OperatorMapResult
                    {
                        FetchOperator = "not-like",
                        Prefix = "%",
                        Suffix = "%"
                    };

                case "beginswith":
                case "startswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Suffix = "%"
                    };

                case "notbeginswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "not-like",
                        Suffix = "%"
                    };

                case "endswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "like",
                        Prefix = "%"
                    };

                case "notendswith":
                    return new OperatorMapResult
                    {
                        FetchOperator = "not-like",
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

                case "between":
                    return new OperatorMapResult
                    {
                        IsBetween = true
                    };

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