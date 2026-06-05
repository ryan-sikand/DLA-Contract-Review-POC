using UiPath.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.CodedWorkflows;

namespace IDP_DD2579
{
    public class ISConnections
    {
        private readonly ICodedWorkflowsServiceContainer _container;

        public ISConnections(ICodedWorkflowsServiceContainer container) => _container = container;

        public DataServiceISConnections DataService => new(_container);
    }

    public class DataServiceISConnections
    {
        private readonly ICodedWorkflowsServiceContainer _container;

        public DataServiceISConnections(ICodedWorkflowsServiceContainer container) => _container = container;

        public ConnectorConnection DefaultConnection =>
            new("ab70558e-d6f2-4d3a-acbd-c1cf62bd4c04", "uipath-uipath-dataservice", _container);
    }
}
