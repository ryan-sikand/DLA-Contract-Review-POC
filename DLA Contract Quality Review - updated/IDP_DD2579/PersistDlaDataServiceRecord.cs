using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using UiPath.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.Models;
using UiPath.IntegrationService.Activities.Runtime.Models.ConnectorMetadata;

namespace IDP_DD2579
{
    public class PersistDlaDataServiceRecord : CodedWorkflow
    {
        [Workflow]
        public async Task Execute(string in_RecordJson)
        {
            if (string.IsNullOrWhiteSpace(in_RecordJson))
            {
                throw new ArgumentException("DLA Data Service payload is empty.", nameof(in_RecordJson));
            }

            using var document = JsonDocument.Parse(in_RecordJson);
            var root = document.RootElement;
            var body = new Dictionary<string, object?>();

            AddString(body, root, "contractBatchId", "ContractBatchId");
            AddString(body, root, "documentRecordKey", "DocumentRecordKey");
            AddString(body, root, "documentType", "DocumentType");
            AddString(body, root, "sourceFilePath", "SourceFilePath");
            AddString(body, root, "sourceFileName", "SourceFileName");
            AddString(body, root, "extractionProjectName", "ExtractionProjectName");
            AddString(body, root, "extractionProjectVersion", "ExtractionProjectVersion");
            AddString(body, root, "validationState", "ValidationState");
            AddString(body, root, "processingState", "ProcessingState");
            body["ExtractedJson"] = in_RecordJson;
            body["AgentInputJson"] = in_RecordJson;
            AddString(body, root, "naicsCode", "NAICSCode");
            AddString(body, root, "productOrServiceCode", "ProductOrServiceCode");
            AddString(body, root, "sf1449Psc", "SF1449PSC");
            AddString(body, root, "sizeStandard", "SizeStandard");
            AddString(body, root, "issuedByCode", "IssuedByCode");
            AddString(body, root, "suppliesServices", "SuppliesServices");
            AddString(body, root, "itemServiceDescription", "ItemServiceDescription");

            var connection = new ISConnections(services.Container).DataService.DefaultConnection;
            if (body.TryGetValue("SourceFileName", out var sourceFileName) &&
                sourceFileName is string fileName &&
                !string.IsNullOrWhiteSpace(fileName) &&
                await RecordExistsAsync(connection, fileName))
            {
                return;
            }

            var configuration = new CodedConnectorConfiguration(
                connection: connection,
                objectName: "CreateEntityRecordCurated",
                operation: Operation.Create,
                httpMethod: "POST",
                path: "/v2/{entityName}/CreateEntityRecord");

            var request = new ConnectorRequest
            {
                PathParameters = new() { ["entityName"] = "DLADataService" },
                QueryParameters = new() { ["expansionLevel"] = "1" },
                BodyParameters = body
            };

            await connection.ExecuteAsync(configuration, request);
        }

        private static async Task<bool> RecordExistsAsync(dynamic connection, string sourceFileName)
        {
            var configuration = new CodedConnectorConfiguration(
                connection: connection,
                objectName: "QueryEntityRecordsCurated",
                operation: Operation.Create,
                httpMethod: "POST",
                path: "/v2/{entityName}/qer");

            var request = new ConnectorRequest
            {
                PathParameters = new() { ["entityName"] = "DLADataService" },
                QueryParameters = new()
                {
                    ["queryExpression"] = $"SourceFileName = '{EscapeCeqlValue(sourceFileName)}'",
                    ["limit"] = "1",
                    ["expansionLevel"] = "1"
                },
                BodyParameters = new Dictionary<string, object?>()
            };

            var response = await connection.ExecuteAsync(configuration, request);
            if (response is null)
            {
                return false;
            }

            using var responseDocument = JsonDocument.Parse(JsonSerializer.Serialize(response));
            return responseDocument.RootElement.ValueKind == JsonValueKind.Array &&
                responseDocument.RootElement.GetArrayLength() > 0;
        }

        private static string EscapeCeqlValue(string value)
        {
            return value.Replace("'", "''", StringComparison.Ordinal);
        }

        private static void AddString(
            IDictionary<string, object?> body,
            JsonElement root,
            string jsonName,
            string fieldName)
        {
            if (!root.TryGetProperty(jsonName, out var value) || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            {
                return;
            }

            body[fieldName] = value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
        }
    }
}
