using Microsoft.Xrm.Sdk;
using System;

namespace Microsoft.Crm.Sdk.Messages
{
    internal class SetProcessStageRequest
    {
        public EntityReference Entity { get; set; }
        public Guid StageId { get; set; }
    }
}