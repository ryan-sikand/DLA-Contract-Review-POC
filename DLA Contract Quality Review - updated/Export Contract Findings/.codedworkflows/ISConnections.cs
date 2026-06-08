using UiPath.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.CodedWorkflows;

namespace ExportContractFindings
{
    public class ISConnections
    {
        private readonly ICodedWorkflowsServiceContainer _container;

        public ISConnections(ICodedWorkflowsServiceContainer container) => _container = container;

        public MicrosoftOneDriveISConnections MicrosoftOneDrive => new(_container);
    }

    public class MicrosoftOneDriveISConnections
    {
        private readonly ICodedWorkflowsServiceContainer _container;

        public MicrosoftOneDriveISConnections(ICodedWorkflowsServiceContainer container) => _container = container;

        public ConnectorConnection DefaultConnection =>
            new("765d339a-c285-47a9-b7cb-65f79d7b73e3", "uipath-microsoft-onedrive", _container);
    }
}
