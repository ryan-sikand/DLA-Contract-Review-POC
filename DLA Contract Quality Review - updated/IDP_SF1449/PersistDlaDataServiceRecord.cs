using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using UiPath.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.CodedWorkflows;
using UiPath.IntegrationService.Activities.Runtime.Models;
using UiPath.IntegrationService.Activities.Runtime.Models.ConnectorMetadata;

namespace IDP_SF1449
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
            NormalizeFileIdentity(body);

            var connection = new ISConnections(services.Container).DataService.DefaultConnection;
            if (body.TryGetValue("SourceFileName", out var sourceFileName) &&
                sourceFileName is string fileName &&
                !string.IsNullOrWhiteSpace(fileName))
            {
                var existingRecordIds = await FindExistingRecordIdsAsync(connection, fileName);
                if (existingRecordIds.Count > 0)
                {
                    await ReplaceRecordAsync(connection, existingRecordIds[0], body);
                    return;
                }
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

        private static async Task<IReadOnlyList<string>> FindExistingRecordIdsAsync(dynamic connection, string sourceFileName)
        {
            var configuration = new CodedConnectorConfiguration(
                connection: connection,
                objectName: "QueryEntityRecordsCurated",
                operation: Operation.List,
                httpMethod: "POST",
                path: "/v2/{entityName}/qer");

            var request = new ConnectorRequest
            {
                PathParameters = new() { ["entityName"] = "DLADataService" },
                QueryParameters = new()
                {
                    ["limit"] = "100",
                    ["expansionLevel"] = "1"
                },
                BodyParameters = new Dictionary<string, object?>
                {
                    ["filterGroup"] = new Dictionary<string, object?>
                    {
                        ["logicalOperator"] = 0,
                        ["queryFilters"] = new object[]
                        {
                            new Dictionary<string, object?>
                            {
                                ["fieldName"] = "SourceFileName",
                                ["operator"] = "=",
                                ["value"] = sourceFileName
                            }
                        }
                    }
                }
            };

            var response = await connection.ExecuteAsync(configuration, request);
            if (response is null)
            {
                return Array.Empty<string>();
            }

            using var responseDocument = JsonDocument.Parse(JsonSerializer.Serialize(response));
            return ExtractRecordIds(responseDocument.RootElement);
        }

        private static async Task ReplaceRecordAsync(
            dynamic connection,
            string recordId,
            IDictionary<string, object?> body)
        {
            var configuration = new CodedConnectorConfiguration(
                connection: connection,
                objectName: "UpdateEntityRecordV2",
                operation: Operation.Replace,
                httpMethod: "PUT",
                path: "/v2/{entityName}/UpdateEntityRecord");

            var request = new ConnectorRequest
            {
                PathParameters = new() { ["entityName"] = "DLADataService" },
                QueryParameters = new()
                {
                    ["recordId"] = recordId,
                    ["expansionLevel"] = "1"
                },
                BodyParameters = new Dictionary<string, object?>(body)
            };

            await connection.ExecuteAsync(configuration, request);
        }

        private static List<string> ExtractRecordIds(JsonElement element)
        {
            var ids = new List<string>();

            if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                {
                    ids.AddRange(ExtractRecordIds(item));
                }

                return ids;
            }

            if (element.ValueKind != JsonValueKind.Object)
            {
                return ids;
            }

            if (element.TryGetProperty("Id", out JsonElement id) && id.ValueKind == JsonValueKind.String)
            {
                var idValue = id.GetString();
                if (!string.IsNullOrWhiteSpace(idValue))
                {
                    ids.Add(idValue);
                }
            }

            if (element.TryGetProperty("Records", out JsonElement records))
            {
                ids.AddRange(ExtractRecordIds(records));
            }

            if (element.TryGetProperty("Data", out JsonElement data))
            {
                ids.AddRange(ExtractRecordIds(data));
            }

            return ids;
        }

        private static void NormalizeFileIdentity(IDictionary<string, object?> body)
        {
            var sourceFileName = TryGetBodyString(body, "SourceFileName");
            if (string.IsNullOrWhiteSpace(sourceFileName))
            {
                sourceFileName = TryGetBodyString(body, "SourceFilePath");
            }

            var normalizedFileName = NormalizeSourceFileName(sourceFileName);
            if (string.IsNullOrWhiteSpace(normalizedFileName))
            {
                return;
            }

            body["SourceFileName"] = normalizedFileName;
            body["DocumentRecordKey"] = normalizedFileName;
        }

        private static string? TryGetBodyString(IDictionary<string, object?> body, string fieldName)
        {
            return body.TryGetValue(fieldName, out var value) ? value as string : null;
        }

        private static string NormalizeSourceFileName(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().Replace('\\', '/');
            var lastSlash = normalized.LastIndexOf('/');
            if (lastSlash >= 0 && lastSlash < normalized.Length - 1)
            {
                normalized = normalized.Substring(lastSlash + 1);
            }

            var lastDot = normalized.LastIndexOf('.');
            if (lastDot > 0)
            {
                normalized = normalized.Substring(0, lastDot);
            }

            return normalized.Trim();
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
