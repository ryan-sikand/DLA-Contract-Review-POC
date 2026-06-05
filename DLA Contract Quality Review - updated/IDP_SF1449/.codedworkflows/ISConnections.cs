using UiPath.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.CodedWorkflows;

namespace IDP_SF1449
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
            new("db8da68a-a9ec-4ae6-9f74-2960708f5812", "uipath-uipath-dataservice", _container);
    }
}
