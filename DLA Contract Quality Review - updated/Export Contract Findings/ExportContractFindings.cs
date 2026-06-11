using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.API;
using UiPath.Core.Activities.Storage;
using UiPath.MicrosoftOffice365.Activities.Api;

namespace ExportContractFindings
{
    public class ExportContractFindings : CodedWorkflow
    {
        private const string DefaultBucketName = "DLA_Contract_Docs";
        private const string DefaultFolderPath = "AMER Presales/Public Sector/DLA Contract Quality Review POC";
        private const string DefaultSharePointFolderUrl = "https://uipath.sharepoint.com/sites/CustomerSuccess-Publicsector/Shared%20Documents/SE%20-%20PubSec/2.%20Demos/DLA%20Contract%20Quality%20Review";
        private const string DefaultOneDriveConnectionId = "27b43e53-db23-4b85-ab3d-4c8693756122";
        private const int CustomerCellLimit = 260;
        private static readonly string[] CustomerColumns =
        {
            "Check",
            "Result",
            "Agent Recommendation",
            "Fields Reviewed",
            "Values Compared",
            "Issue",
            "Recommended Action"
        };

        private static readonly string[] RequiredChecks =
        {
            "NAICS / PSC / Size Standard Match",
            "NAICS / SBA Size Standard",
            "Semantic Alignment",
            "SAM Exclusion Search Date",
            "D&F Requirement"
        };

        [Workflow]
        public async Task<(string bucketFilePath, string sharePointUploadStatus, string localWorkbookPath)> Execute(
            string excelReportRowsJson,
            string agentContentJson = "",
            string storageBucketName = "",
            string orchestratorFolderPath = "",
            string sharePointFolderUrl = "")
        {
            if (string.IsNullOrWhiteSpace(excelReportRowsJson))
            {
                throw new ArgumentException("excelReportRowsJson is empty; the agent did not return Excel-ready rows.", nameof(excelReportRowsJson));
            }

            storageBucketName = string.IsNullOrWhiteSpace(storageBucketName) ? DefaultBucketName : storageBucketName;
            orchestratorFolderPath = string.IsNullOrWhiteSpace(orchestratorFolderPath) ? DefaultFolderPath : orchestratorFolderPath;
            sharePointFolderUrl = string.IsNullOrWhiteSpace(sharePointFolderUrl) ? DefaultSharePointFolderUrl : sharePointFolderUrl;

            var rows = ParseRows(excelReportRowsJson);
            var content = ParseObject(agentContentJson);
            var timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            var fileName = $"DLA_Contract_Quality_Findings_{timestamp}.xlsx";
            var localWorkbookPath = Path.Combine(Path.GetTempPath(), fileName);
            var bucketPath = $"reports/{fileName}";

            WriteWorkbook(localWorkbookPath, rows, content, agentContentJson);

            if (!system.PathExists(localWorkbookPath, out var fileResource))
            {
                throw new FileNotFoundException("The generated Excel workbook was not found.", localWorkbookPath);
            }

            system.UploadStorageFile(bucketPath, fileResource, storageBucketName, orchestratorFolderPath, 60000);

            var sharePointStatus = await TryUploadToSharePointAsync(localWorkbookPath, fileName, sharePointFolderUrl);
            Log($"Exported workbook to bucket path '{bucketPath}'. SharePoint status: {sharePointStatus}");

            return (bucketPath, sharePointStatus, localWorkbookPath);
        }

        private async Task<string> TryUploadToSharePointAsync(string localFilePath, string fileName, string sharePointFolderUrl)
        {
            if (string.IsNullOrWhiteSpace(sharePointFolderUrl))
            {
                return "Skipped: sharePointFolderUrl was empty.";
            }

            try
            {
                await Task.Yield();
                var oneDrive = office365.OneDrive(new OneDriveConnection(DefaultOneDriveConnectionId, serviceContainer));
                var destination = oneDrive.GetFolder(sharePointFolderUrl);
                var uploaded = oneDrive.UploadFile(localFilePath, destination, ConflictBehavior.Replace, null);
                return $"Uploaded to SharePoint folder: {sharePointFolderUrl}; item: {uploaded.Name}";
            }
            catch (Exception ex)
            {
                var message = $"SharePoint upload failed: {ex.GetType().Name}: {ex.Message}";
                Log(message);
                throw new InvalidOperationException(message, ex);
            }
        }

        private static List<Dictionary<string, string>> ParseRows(string rowsJson)
        {
            var rows = new List<Dictionary<string, string>>();
            using var document = JsonDocument.Parse(rowsJson);
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                throw new ArgumentException("excelReportRowsJson must be a JSON array.");
            }

            foreach (var item in document.RootElement.EnumerateArray())
            {
                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (item.ValueKind == JsonValueKind.Object)
                {
                    foreach (var property in item.EnumerateObject())
                    {
                        row[property.Name] = ElementToString(property.Value);
                    }
                }

                rows.Add(row);
            }

            return rows;
        }

        private static Dictionary<string, string> ParseObject(string json)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(json))
            {
                return values;
            }

            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    return values;
                }

                foreach (var property in document.RootElement.EnumerateObject())
                {
                    if (property.Value.ValueKind is JsonValueKind.Array or JsonValueKind.Object)
                    {
                        continue;
                    }

                    values[property.Name] = ElementToString(property.Value);
                }
            }
            catch (JsonException)
            {
                values["notes"] = json;
            }

            return values;
        }

        private static string ElementToString(JsonElement element) =>
            element.ValueKind switch
            {
                JsonValueKind.String => element.GetString() ?? "",
                JsonValueKind.Number => element.GetRawText(),
                JsonValueKind.True => "TRUE",
                JsonValueKind.False => "FALSE",
                JsonValueKind.Null => "",
                JsonValueKind.Undefined => "",
                _ => element.GetRawText()
            };

        private static void WriteWorkbook(
            string path,
            IReadOnlyList<Dictionary<string, string>> findings,
            IReadOnlyDictionary<string, string> summary,
            string agentContentJson)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var reportRows = BuildCustomerRows(findings);
            var sourceDocuments = SourceDocumentRows(findings, agentContentJson);
            var technicalRows = TechnicalDetailRows(agentContentJson, findings, reportRows);
            var customerGrid = CustomerSummaryGrid(summary, reportRows, sourceDocuments.Count);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
            AddText(archive, "[Content_Types].xml", ContentTypesXml());
            AddText(archive, "_rels/.rels", RootRelationshipsXml());
            AddText(archive, "xl/workbook.xml", WorkbookXml());
            AddText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml());
            AddText(archive, "xl/styles.xml", StylesXml());
            AddText(archive, "xl/worksheets/sheet1.xml", SheetGridXml("Contract Quality Review Results", customerGrid));
            AddText(archive, "xl/worksheets/sheet2.xml", SheetXml("Technical Details", technicalRows));
            AddText(archive, "xl/worksheets/sheet3.xml", SheetXml("Source Documents", sourceDocuments));
        }

        private static IReadOnlyList<IReadOnlyList<string>> CustomerSummaryGrid(
            IReadOnlyDictionary<string, string> summary,
            IReadOnlyList<Dictionary<string, string>> findings,
            int sourceDocumentCount)
        {
            var passCount = findings.Count(r => Get(r, "Result").Equals("Pass", StringComparison.OrdinalIgnoreCase));
            var flagCount = findings.Count(r => Get(r, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase));
            var notApplicableCount = findings.Count(r => Get(r, "Result").Equals("Not Applicable", StringComparison.OrdinalIgnoreCase));
            var overallStatus = FirstNonEmpty(Get(summary, "overall_status"), CustomerOverallStatus(findings));
            var contractPackage = FirstNonEmpty(Get(summary, "contract_package"), Get(summary, "contract_batch_id"), Get(summary, "contractBatchId"), "Not specified");
            var created = FirstNonEmpty(Get(summary, "created"), DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'"));
            var documentsProcessed = FirstNonEmpty(Get(summary, "documents_processed"), sourceDocumentCount > 0 ? sourceDocumentCount.ToString() : "");
            var fieldsReviewed = FirstNonEmpty(Get(summary, "fields_reviewed"), $"{CountReviewedFields(findings)} field groups");
            var rulesEvaluated = FirstNonEmpty(Get(summary, "business_rules_evaluated"), findings.Count.ToString());
            var resultSummary = FirstNonEmpty(Get(summary, "result_summary"), $"{passCount} Pass | {flagCount} Flag | {notApplicableCount} Not Applicable");

            var rows = new List<IReadOnlyList<string>>
            {
                new[] { "Contract Quality Review Results" },
                Array.Empty<string>(),
                new[] { "Overall Status", overallStatus },
                new[] { "Contract Package", contractPackage },
                new[] { "Created", created },
                new[] { "Documents Processed", documentsProcessed },
                new[] { "Fields Reviewed", fieldsReviewed },
                new[] { "Business Rules Evaluated", rulesEvaluated },
                new[] { "Result Summary", resultSummary },
                Array.Empty<string>(),
                new[] { "Business Rule Results" },
                CustomerColumns
            };

            rows.AddRange(findings.Select(row => CustomerColumns.Select(column => Get(row, column)).ToArray()));

            var completenessNote = DataCompletenessNote(findings);
            if (!string.IsNullOrWhiteSpace(completenessNote))
            {
                rows.Add(Array.Empty<string>());
                rows.Add(new[] { "Data Completeness Notes" });
                rows.Add(new[] { completenessNote });
            }

            return rows;
        }

        private static List<Dictionary<string, string>> BuildCustomerRows(IReadOnlyList<Dictionary<string, string>> findings)
        {
            var candidates = findings.Select(ToCustomerCandidate)
                .Where(row => !string.IsNullOrWhiteSpace(Get(row, "Check")))
                .ToList();

            var rows = new List<Dictionary<string, string>>();
            foreach (var requiredCheck in RequiredChecks)
            {
                var matchingRows = candidates
                    .Where(row => Get(row, "Check").Equals(requiredCheck, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                rows.Add(matchingRows.Count == 0
                    ? MissingCheckRow(requiredCheck)
                    : MergeCustomerRows(requiredCheck, matchingRows));
            }

            return rows;
        }

        private static Dictionary<string, string> ToCustomerCandidate(IReadOnlyDictionary<string, string> row)
        {
            var checkId = FirstNonEmpty(Get(row, "CheckId"), Get(row, "Check"));
            var checkName = FirstNonEmpty(Get(row, "CheckName"), Get(row, "Review Item"));
            var check = MapCustomerCheck(FirstNonEmpty(Get(row, "Check"), checkName, checkId));
            var result = NormalizeResult(Get(row, "Result"));
            var issue = FirstNonEmpty(Get(row, "Issue"), Get(row, "Exception"), Get(row, "What We Found"), Get(row, "Evidence"));
            var action = FirstNonEmpty(Get(row, "Recommended Action"), Get(row, "RecommendedAction"));

            return new Dictionary<string, string>
            {
                ["Check"] = check,
                ["Result"] = result,
                ["Agent Recommendation"] = NormalizeRecommendation(FirstNonEmpty(Get(row, "Agent Recommendation"), Get(row, "Recommendation")), result, issue),
                ["Fields Reviewed"] = FirstNonEmpty(Get(row, "Fields Reviewed"), Get(row, "Field Reviewed"), Get(row, "FieldName"), DefaultFieldsReviewed(check)),
                ["Values Compared"] = FirstNonEmpty(Get(row, "Values Compared"), Get(row, "ComparedValues"), CombinedValues(row), DefaultValuesCompared(check)),
                ["Issue"] = NormalizeIssue(result, issue),
                ["Recommended Action"] = NormalizeAction(result, action, check)
            };
        }

        private static Dictionary<string, string> MergeCustomerRows(string check, IReadOnlyList<Dictionary<string, string>> rows)
        {
            var result = rows.Any(r => Get(r, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase))
                ? "Flag"
                : rows.Any(r => Get(r, "Result").Equals("Pass", StringComparison.OrdinalIgnoreCase))
                    ? "Pass"
                    : "Not Applicable";

            var issue = result.Equals("Pass", StringComparison.OrdinalIgnoreCase)
                ? "No issue identified"
                : JoinDistinct(rows.Select(r => Get(r, "Issue")));

            return new Dictionary<string, string>
            {
                ["Check"] = check,
                ["Result"] = result,
                ["Agent Recommendation"] = NormalizeRecommendation(JoinDistinct(rows.Select(r => Get(r, "Agent Recommendation"))), result, issue),
                ["Fields Reviewed"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Fields Reviewed"))), DefaultFieldsReviewed(check))),
                ["Values Compared"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Values Compared"))), DefaultValuesCompared(check))),
                ["Issue"] = CleanCustomerText(NormalizeIssue(result, issue)),
                ["Recommended Action"] = CleanCustomerText(NormalizeAction(result, JoinDistinct(rows.Select(r => Get(r, "Recommended Action"))), check))
            };
        }

        private static Dictionary<string, string> MissingCheckRow(string check) =>
            new()
            {
                ["Check"] = check,
                ["Result"] = "Flag",
                ["Agent Recommendation"] = "Missing evidence",
                ["Fields Reviewed"] = DefaultFieldsReviewed(check),
                ["Values Compared"] = DefaultValuesCompared(check),
                ["Issue"] = "The agent did not return enough evidence to complete this check.",
                ["Recommended Action"] = DefaultRecommendedAction(check)
            };

        private static IReadOnlyList<Dictionary<string, string>> SourceDocumentRows(IReadOnlyList<Dictionary<string, string>> findings, string agentContentJson)
        {
            var rows = new List<Dictionary<string, string>>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddSourceDocumentsFromContent(agentContentJson, rows, seen);

            foreach (var finding in findings)
            {
                var files = SplitList(FirstNonEmpty(Get(finding, "SourceFileName"), Get(finding, "Documents Reviewed")));
                var types = SplitList(FirstNonEmpty(Get(finding, "DocumentType"), Get(finding, "Document Type")));
                if (files.Count == 0 && types.Count == 0)
                {
                    continue;
                }

                var count = Math.Max(files.Count, types.Count);
                for (var i = 0; i < count; i++)
                {
                    var file = i < files.Count ? files[i] : "";
                    var type = i < types.Count ? types[i] : FirstNonEmpty(Get(finding, "DocumentType"), Get(finding, "Document Type"), InferDocumentType(file));
                    var key = $"{type}|{file}";
                    if (!seen.Add(key))
                    {
                        continue;
                    }

                    rows.Add(new Dictionary<string, string>
                    {
                        ["Document Type"] = type,
                        ["File Name"] = file,
                        ["Used In Check"] = FirstNonEmpty(Get(finding, "CheckId"), Get(finding, "Check"), Get(finding, "CheckName")),
                        ["Notes"] = FirstNonEmpty(Get(finding, "Issue"), Get(finding, "Exception"), Get(finding, "Evidence"))
                    });
                }
            }

            return rows;
        }

        private static void AddSourceDocumentsFromContent(string agentContentJson, List<Dictionary<string, string>> rows, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(agentContentJson))
            {
                return;
            }

            try
            {
                using var document = JsonDocument.Parse(agentContentJson);
                if (TryGetProperty(document.RootElement, "source_records", out var sourceRecords) && sourceRecords.ValueKind == JsonValueKind.Array)
                {
                    foreach (var record in sourceRecords.EnumerateArray())
                    {
                        AddSourceDocument(record, rows, seen);
                    }
                }

                if (TryGetProperty(document.RootElement, "technical_detail", out var technicalDetail) &&
                    TryGetProperty(technicalDetail, "source_records", out var technicalSourceRecords) &&
                    technicalSourceRecords.ValueKind == JsonValueKind.Array)
                {
                    foreach (var record in technicalSourceRecords.EnumerateArray())
                    {
                        AddSourceDocument(record, rows, seen);
                    }
                }
            }
            catch (JsonException)
            {
                return;
            }
        }

        private static void AddSourceDocument(JsonElement record, List<Dictionary<string, string>> rows, HashSet<string> seen)
        {
            if (record.ValueKind != JsonValueKind.Object)
            {
                return;
            }

            var file = FirstNonEmpty(
                GetJsonString(record, "sourceFileName"),
                GetJsonString(record, "source_file_name"),
                GetJsonString(record, "sourceFilePath"),
                GetJsonString(record, "source_file_path"),
                GetJsonString(record, "fileName"),
                GetJsonString(record, "filename"));
            var type = FirstNonEmpty(
                GetJsonString(record, "documentType"),
                GetJsonString(record, "document_type"),
                InferDocumentType(file));
            var key = $"{type}|{file}";
            if (string.IsNullOrWhiteSpace(key) || !seen.Add(key))
            {
                return;
            }

            rows.Add(new Dictionary<string, string>
            {
                ["Document Type"] = type,
                ["File Name"] = file,
                ["Used In Check"] = "Technical evidence",
                ["Notes"] = FirstNonEmpty(
                    GetJsonString(record, "processingState"),
                    GetJsonString(record, "validationState"),
                    GetJsonString(record, "extractionProjectName"))
            });
        }

        private static IReadOnlyList<Dictionary<string, string>> TechnicalDetailRows(
            string agentContentJson,
            IReadOnlyList<Dictionary<string, string>> originalFindings,
            IReadOnlyList<Dictionary<string, string>> customerRows)
        {
            var rows = new List<Dictionary<string, string>>
            {
                TechnicalRow("Report", "Generated", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'")),
                TechnicalRow("Report", "Customer-facing sheet", "Contract Quality Review Results")
            };

            if (!string.IsNullOrWhiteSpace(agentContentJson))
            {
                try
                {
                    using var document = JsonDocument.Parse(agentContentJson);
                    FlattenJson(rows, "Agent Content", "", document.RootElement, 0);
                }
                catch (JsonException)
                {
                    rows.Add(TechnicalRow("Agent Content", "Raw", Truncate(agentContentJson, 4000)));
                }
            }

            AddTabularTechnicalRows(rows, "Original Agent Row", originalFindings);
            AddTabularTechnicalRows(rows, "Customer Summary Row", customerRows);
            return rows;
        }

        private static void AddTabularTechnicalRows(List<Dictionary<string, string>> rows, string section, IReadOnlyList<Dictionary<string, string>> sourceRows)
        {
            for (var i = 0; i < sourceRows.Count; i++)
            {
                foreach (var pair in sourceRows[i])
                {
                    rows.Add(TechnicalRow($"{section} {i + 1}", pair.Key, pair.Value));
                }
            }
        }

        private static Dictionary<string, string> TechnicalRow(string section, string field, string value) =>
            new()
            {
                ["Section"] = section,
                ["Field"] = field,
                ["Value"] = Truncate(value, 4000)
            };

        private static void FlattenJson(List<Dictionary<string, string>> rows, string section, string path, JsonElement element, int depth)
        {
            if (depth > 5)
            {
                rows.Add(TechnicalRow(section, path, element.GetRawText()));
                return;
            }

            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var property in element.EnumerateObject())
                    {
                        var childPath = string.IsNullOrWhiteSpace(path) ? property.Name : $"{path}.{property.Name}";
                        FlattenJson(rows, section, childPath, property.Value, depth + 1);
                    }
                    break;
                case JsonValueKind.Array:
                    var index = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        if (index >= 100)
                        {
                            rows.Add(TechnicalRow(section, $"{path}[more]", "Additional values omitted from the technical workbook tab."));
                            break;
                        }

                        FlattenJson(rows, section, $"{path}[{index}]", item, depth + 1);
                        index++;
                    }

                    if (index == 0)
                    {
                        rows.Add(TechnicalRow(section, path, ""));
                    }
                    break;
                default:
                    rows.Add(TechnicalRow(section, path, ElementToString(element)));
                    break;
            }
        }

        private static string InferDocumentType(string fileName)
        {
            var name = (fileName ?? "").ToUpperInvariant();
            if (name.Contains("DD2579", StringComparison.Ordinal))
            {
                return "DD2579";
            }

            if (name.Contains("SF1449", StringComparison.Ordinal))
            {
                return "SF1449";
            }

            if (name.Contains("SAAD", StringComparison.Ordinal))
            {
                return "SAAD";
            }

            if (name.Contains("DF", StringComparison.Ordinal) || name.Contains("D&F", StringComparison.Ordinal))
            {
                return "D&F";
            }

            return "";
        }

        private static string MapCustomerCheck(string value)
        {
            var text = (value ?? "").Trim();
            var normalized = text.ToUpperInvariant();
            if (normalized.Contains("1A", StringComparison.Ordinal) ||
                normalized.Contains("THREE-WAY", StringComparison.Ordinal) ||
                normalized.Contains("NAICS/PSC", StringComparison.Ordinal) ||
                (normalized.Contains("NAICS", StringComparison.Ordinal) &&
                 normalized.Contains("PSC", StringComparison.Ordinal) &&
                 normalized.Contains("SIZE", StringComparison.Ordinal)))
            {
                return "NAICS / PSC / Size Standard Match";
            }

            if (normalized.Contains("1B", StringComparison.Ordinal) ||
                normalized.Contains("SBA", StringComparison.Ordinal) ||
                normalized.Contains("SIZE STANDARD CONSISTENCY", StringComparison.Ordinal))
            {
                return "NAICS / SBA Size Standard";
            }

            if (normalized.Contains("1C", StringComparison.Ordinal) ||
                normalized.Contains("SEMANTIC", StringComparison.Ordinal))
            {
                return "Semantic Alignment";
            }

            if (normalized.Contains("2A", StringComparison.Ordinal) ||
                normalized.Contains("SAM", StringComparison.Ordinal))
            {
                return "SAM Exclusion Search Date";
            }

            if (normalized.Contains("2B", StringComparison.Ordinal) ||
                normalized.Contains("D&F", StringComparison.Ordinal) ||
                normalized.Contains("D\\u0026F", StringComparison.Ordinal) ||
                normalized.Contains("DETERMINATION", StringComparison.Ordinal))
            {
                return "D&F Requirement";
            }

            return text;
        }

        private static string NormalizeResult(string value)
        {
            var text = (value ?? "").Trim();
            var normalized = text.ToUpperInvariant();
            if (normalized is "PASS" or "PASSED")
            {
                return "Pass";
            }

            if (normalized is "N/A" or "NA" or "NOT APPLICABLE" or "NOT_APPLICABLE")
            {
                return "Not Applicable";
            }

            if (normalized is "FLAG" or "FLAGGED" or "REVIEW" or "WARNING" or "FAIL" or "FAILED")
            {
                return "Flag";
            }

            if (normalized.Contains("PASS", StringComparison.Ordinal))
            {
                return "Pass";
            }

            if (normalized.Contains("NOT APPLICABLE", StringComparison.Ordinal))
            {
                return "Not Applicable";
            }

            return string.IsNullOrWhiteSpace(text) ? "Flag" : "Flag";
        }

        private static string NormalizeRecommendation(string value, string result, string issue)
        {
            if (result.Equals("Pass", StringComparison.OrdinalIgnoreCase))
            {
                return "No action needed";
            }

            if (result.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
            {
                return "Not applicable";
            }

            var text = (value ?? "").Trim();
            var normalized = text.ToUpperInvariant();
            if (normalized.Contains("MISSING", StringComparison.Ordinal) || IndicatesMissingEvidence(issue))
            {
                return "Missing evidence";
            }

            if (normalized.Contains("CONFIRM", StringComparison.Ordinal))
            {
                return "Confirm";
            }

            return "Review";
        }

        private static string DefaultFieldsReviewed(string check) =>
            check switch
            {
                "NAICS / PSC / Size Standard Match" => "NAICS, PSC, size standard",
                "NAICS / SBA Size Standard" => "NAICS, size standard",
                "Semantic Alignment" => "Item/service description, PSC, NAICS",
                "SAM Exclusion Search Date" => "SAM exclusion search date, award date, issued-by code",
                "D&F Requirement" => "CLIN type, award date, D&F presence/signature date",
                _ => ""
            };

        private static string DefaultValuesCompared(string check) =>
            check switch
            {
                "NAICS / PSC / Size Standard Match" => "DD2579 vs. SF1449 solicitation vs. SF1449 award",
                "NAICS / SBA Size Standard" => "DD2579 NAICS and size standard vs. SBA Table of Size Standards",
                "Semantic Alignment" => "DD2579 item/service description vs. identified PSC and NAICS",
                "SAM Exclusion Search Date" => "SAAD search date vs. SF1449 award date and required 4- or 7-business-day window",
                "D&F Requirement" => "SF1449 CLIN detail vs. D&F requirement and D&F signature date",
                _ => ""
            };

        private static string NormalizeIssue(string result, string issue)
        {
            if (result.Equals("Pass", StringComparison.OrdinalIgnoreCase))
            {
                return "No issue identified";
            }

            if (result.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(issue))
            {
                return "This check does not apply to this package.";
            }

            return CleanCustomerText(FirstNonEmpty(issue, "Reviewer confirmation is needed."));
        }

        private static string NormalizeAction(string result, string action, string check)
        {
            if (result.Equals("Pass", StringComparison.OrdinalIgnoreCase))
            {
                return "No action needed";
            }

            if (result.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(action))
            {
                return "No action needed.";
            }

            return CleanCustomerText(FirstNonEmpty(action, DefaultRecommendedAction(check)));
        }

        private static string DefaultRecommendedAction(string check) =>
            check switch
            {
                "NAICS / PSC / Size Standard Match" => "Confirm conflicting or missing values in the DD2579 and SF1449 documents, including Block 20 and continuation pages for PSC.",
                "NAICS / SBA Size Standard" => "Confirm the DD2579 NAICS and size standard against the SBA Table of Size Standards.",
                "Semantic Alignment" => "Review whether the PSC and NAICS classifications align with the item/service description.",
                "SAM Exclusion Search Date" => "Confirm the SAAD search date, award date, issued-by code, and required business-day window.",
                "D&F Requirement" => "Review award schedule/continuation pages. If T&M or labor-hour CLINs are present, confirm a signed D&F predates the award.",
                _ => "Review the supporting source documents."
            };

        private static string CleanCustomerText(string value)
        {
            var text = (value ?? "").Replace("\r", " ", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal).Trim();
            while (text.Contains("  ", StringComparison.Ordinal))
            {
                text = text.Replace("  ", " ", StringComparison.Ordinal);
            }

            if ((text.StartsWith("{", StringComparison.Ordinal) || text.StartsWith("[", StringComparison.Ordinal)) && text.Length > 80)
            {
                return "Detailed evidence is available in the Technical Details tab.";
            }

            var upper = text.ToUpperInvariant();
            if (upper.Contains("STACK TRACE", StringComparison.Ordinal) ||
                upper.Contains("QUEUE ITEM", StringComparison.Ordinal) ||
                upper.Contains("ROBOT ", StringComparison.Ordinal) ||
                upper.Contains("JOB ID", StringComparison.Ordinal))
            {
                return "Detailed processing information is available in the Technical Details tab.";
            }

            return Truncate(text, CustomerCellLimit);
        }

        private static string JoinDistinct(IEnumerable<string> values) =>
            string.Join("; ", values
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => v.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase));

        private static string DataCompletenessNote(IReadOnlyList<Dictionary<string, string>> findings)
        {
            if (!findings.Any(row =>
                    Get(row, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase) &&
                    (Get(row, "Agent Recommendation").Equals("Missing evidence", StringComparison.OrdinalIgnoreCase) ||
                     IndicatesMissingEvidence(Get(row, "Issue")))))
            {
                return "";
            }

            return "Some checks were flagged because the available documents did not fully confirm the required evidence. A flag does not necessarily mean the contract is non-compliant; it means reviewer confirmation is recommended.";
        }

        private static bool IndicatesMissingEvidence(string value)
        {
            var normalized = (value ?? "").ToUpperInvariant();
            return normalized.Contains("MISSING", StringComparison.Ordinal) ||
                   normalized.Contains("CANNOT", StringComparison.Ordinal) ||
                   normalized.Contains("COULD NOT", StringComparison.Ordinal) ||
                   normalized.Contains("NOT FOUND", StringComparison.Ordinal) ||
                   normalized.Contains("NOT ENOUGH", StringComparison.Ordinal) ||
                   normalized.Contains("INSUFFICIENT", StringComparison.Ordinal) ||
                   normalized.Contains("UNCONFIRMED", StringComparison.Ordinal) ||
                   normalized.Contains("NOT RETURN", StringComparison.Ordinal);
        }

        private static string CustomerOverallStatus(IReadOnlyList<Dictionary<string, string>> findings)
        {
            if (findings.Any(row => Get(row, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase)))
            {
                return "Reviewer Action Recommended";
            }

            return findings.Any(row => Get(row, "Result").Equals("Pass", StringComparison.OrdinalIgnoreCase))
                ? "No Issues Identified"
                : "No Applicable Checks";
        }

        private static int CountReviewedFields(IReadOnlyList<Dictionary<string, string>> findings) =>
            findings.SelectMany(row => SplitList(Get(row, "Fields Reviewed")))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();

        private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var property in element.EnumerateObject())
                {
                    if (property.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        value = property.Value;
                        return true;
                    }
                }
            }

            value = default;
            return false;
        }

        private static string GetJsonString(JsonElement element, string name) =>
            TryGetProperty(element, name, out var value) ? ElementToString(value) : "";

        private static string Truncate(string value, int maxLength)
        {
            var text = value ?? "";
            return text.Length <= maxLength ? text : text.Substring(0, Math.Max(0, maxLength - 3)) + "...";
        }

        private static string SheetGridXml(string sheetName, IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<cols>");
            var widths = new[] { 34, 18, 24, 34, 48, 48, 58 };
            for (var i = 0; i < widths.Length; i++)
            {
                sb.Append("<col min=\"").Append(i + 1).Append("\" max=\"").Append(i + 1).Append("\" width=\"").Append(widths[i]).Append("\" customWidth=\"1\"/>");
            }
            sb.Append("</cols><sheetData>");

            for (var i = 0; i < rows.Count; i++)
            {
                WriteGridRow(sb, i + 1, rows[i]);
            }

            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        private static void WriteGridRow(StringBuilder sb, int rowIndex, IReadOnlyList<string> values)
        {
            sb.Append("<row r=\"").Append(rowIndex).Append("\">");
            for (var i = 0; i < values.Count; i++)
            {
                var cellRef = ColumnName(i + 1) + rowIndex;
                sb.Append("<c r=\"").Append(cellRef).Append("\" t=\"inlineStr\"><is><t>");
                sb.Append(XmlEscape(values[i]));
                sb.Append("</t></is></c>");
            }

            sb.Append("</row>");
        }

        private static string SheetXml(string sheetName, IReadOnlyList<Dictionary<string, string>> rows)
        {
            var headers = rows.SelectMany(r => r.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .DefaultIfEmpty(sheetName)
                .ToList();

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            WriteRow(sb, 1, headers.ToDictionary(h => h, h => h));

            for (var i = 0; i < rows.Count; i++)
            {
                var current = rows[i];
                WriteRow(sb, i + 2, headers.ToDictionary(h => h, h => Get(current, h)));
            }

            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        private static void WriteRow(StringBuilder sb, int rowIndex, Dictionary<string, string> values)
        {
            sb.Append("<row r=\"").Append(rowIndex).Append("\">");
            var column = 1;
            foreach (var value in values.Values)
            {
                var cellRef = ColumnName(column) + rowIndex;
                sb.Append("<c r=\"").Append(cellRef).Append("\" t=\"inlineStr\"><is><t>");
                sb.Append(XmlEscape(value));
                sb.Append("</t></is></c>");
                column++;
            }

            sb.Append("</row>");
        }

        private static string Get(IReadOnlyDictionary<string, string> row, string key) =>
            row.TryGetValue(key, out var value) ? value ?? "" : "";

        private static string FirstNonEmpty(params string[] values) =>
            values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";

        private static string CombinedValues(IReadOnlyDictionary<string, string> row)
        {
            var values = new[]
            {
                ("DD2579", Get(row, "DD2579Value")),
                ("SF1449", Get(row, "SF1449Value")),
                ("SAAD", Get(row, "SAADValue")),
                ("D&F", Get(row, "DFTRValue"))
            }
            .Where(v => !string.IsNullOrWhiteSpace(v.Item2))
            .Select(v => $"{v.Item1}: {v.Item2}");

            return string.Join(" | ", values);
        }

        private static List<string> SplitList(string value) =>
            (value ?? "")
                .Split(new[] { ';', '|', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(v => v.Trim())
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .ToList();

        private static string FriendlyDataSource(string value)
        {
            var normalized = (value ?? "").Trim();
            return normalized.ToUpperInvariant() switch
            {
                "IDP_JSON_PRIMARY" => "IDP extraction JSON",
                "DATA_FABRIC_FALLBACK" => "Data Fabric fallback",
                "MIXED_JSON_AND_DATA_FABRIC" => "IDP extraction JSON and Data Fabric",
                _ => normalized
            };
        }

        private static string FriendlyBoolean(string value)
        {
            if (value.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                return "Yes";
            }

            if (value.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                return "No";
            }

            return value;
        }

        private static string HighestResult(IReadOnlyList<Dictionary<string, string>> findings)
        {
            if (findings.Any(r => Get(r, "Result").Equals("FLAG", StringComparison.OrdinalIgnoreCase)))
            {
                return "FLAG";
            }

            if (findings.Any(r => Get(r, "Result").Equals("REVIEW", StringComparison.OrdinalIgnoreCase)))
            {
                return "REVIEW";
            }

            if (findings.Any(r => Get(r, "Result").Equals("PASS", StringComparison.OrdinalIgnoreCase)))
            {
                return "PASS";
            }

            return findings.Any() ? "N/A" : "";
        }

        private static string ColumnName(int index)
        {
            var name = "";
            while (index > 0)
            {
                index--;
                name = (char)('A' + index % 26) + name;
                index /= 26;
            }

            return name;
        }

        private static string XmlEscape(string value) =>
            (value ?? "")
                .Replace("&", "&amp;", StringComparison.Ordinal)
                .Replace("<", "&lt;", StringComparison.Ordinal)
                .Replace(">", "&gt;", StringComparison.Ordinal)
                .Replace("\"", "&quot;", StringComparison.Ordinal)
                .Replace("'", "&apos;", StringComparison.Ordinal);

        private static void AddText(ZipArchive archive, string entryName, string text)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Encoding.UTF8);
            writer.Write(text);
        }

        private static string ContentTypesXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "</Types>";

        private static string RootRelationshipsXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>";

        private static string WorkbookXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets>" +
            "<sheet name=\"Contract Quality Review Results\" sheetId=\"1\" r:id=\"rId1\"/>" +
            "<sheet name=\"Technical Details\" sheetId=\"2\" r:id=\"rId2\"/>" +
            "<sheet name=\"Source Documents\" sheetId=\"3\" r:id=\"rId3\"/>" +
            "</sheets></workbook>";

        private static string WorkbookRelationshipsXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" +
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>" +
            "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>";

        private static string StylesXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
            "</styleSheet>";
    }
}
