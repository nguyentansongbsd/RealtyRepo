using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Action_MatchUnit_Approve
{
    public class Action_MatchUnit_Approve : IPlugin
    {
        private IPluginExecutionContext _context;
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private IOrganizationServiceFactory _serviceFactory;

        public void Execute(IServiceProvider serviceProvider)
        {
            this._context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            this._tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            this._serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            this._service = _serviceFactory.CreateOrganizationService(this._context.UserId);

        }
    }
}
