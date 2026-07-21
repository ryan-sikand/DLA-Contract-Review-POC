using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
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
        private const int CustomerCellLimit = 1800;
        private static readonly string[] CustomerColumns =
        {
            "Check",
            "Result",
            "Agent Recommendation",
            "Fields Reviewed",
            "Values Compared",
            "Result Summary",
            "Recommended Action",
            "Data Completeness Notes",
            "Source Documents"
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
        public async Task<(string bucketFilePath, string sharePointUploadStatus, string localWorkbookPath, string normalizedExcelReportRowsJson)> Execute(
            string excelReportRowsJson,
            string agentContentJson = "",
            string storageBucketName = "",
            string orchestratorFolderPath = "",
            string sharePointFolderUrl = "",
            string contractDataJson = "")
        {
            if (string.IsNullOrWhiteSpace(excelReportRowsJson))
            {
                throw new ArgumentException("excelReportRowsJson is empty; the agent did not return Excel-ready rows.", nameof(excelReportRowsJson));
            }

            storageBucketName = string.IsNullOrWhiteSpace(storageBucketName) ? DefaultBucketName : storageBucketName;
            orchestratorFolderPath = string.IsNullOrWhiteSpace(orchestratorFolderPath) ? DefaultFolderPath : orchestratorFolderPath;
            sharePointFolderUrl = string.IsNullOrWhiteSpace(sharePointFolderUrl) ? DefaultSharePointFolderUrl : sharePointFolderUrl;

            var rows = ParseRows(excelReportRowsJson);
            ApplyDeterministicSamValidation(rows, contractDataJson);
            ApplySamResultConsistencyGuard(rows);
            ApplySemanticAlignmentConsistencyGuard(rows, contractDataJson);
            var normalizedExcelReportRowsJson = SerializeRows(rows);
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

            return (bucketPath, sharePointStatus, localWorkbookPath, normalizedExcelReportRowsJson);
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

        private static void ApplyDeterministicSamValidation(List<Dictionary<string, string>> rows, string contractDataJson)
        {
            if (!TryBuildSamValidationRow(contractDataJson, out var samRow))
            {
                ApplySamRowConsistencyFallback(rows);
                return;
            }

            var existing = rows.FirstOrDefault(row =>
                MapCustomerCheck(FirstNonEmpty(Get(row, "Check"), Get(row, "CheckName"), Get(row, "Review Item")))
                    .Equals("SAM Exclusion Search Date", StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                rows.Add(samRow);
                return;
            }

            existing["Check"] = "SAM Exclusion Search Date";
            foreach (var pair in samRow)
            {
                existing[pair.Key] = pair.Value;
            }
        }

        private static void ApplySamRowConsistencyFallback(List<Dictionary<string, string>> rows)
        {
            var existing = rows.FirstOrDefault(row =>
                MapCustomerCheck(FirstNonEmpty(Get(row, "Check"), Get(row, "CheckName"), Get(row, "Review Item")))
                    .Equals("SAM Exclusion Search Date", StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                return;
            }

            var evidence = string.Join(" ",
                Get(existing, "Values Compared"),
                Get(existing, "Result Summary"),
                Get(existing, "Data Completeness Notes"));

            if (!evidence.Contains("SAAD", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var hasDates = TryExtractSamDatesAndWindow(
                evidence,
                out var samDate,
                out var awardDate,
                out var requiredWindow,
                out var issuedByCode,
                out var samFile,
                out var awardFile);
            int count;
            bool isAfterAward;
            string countedDatesText;
            if (hasDates)
            {
                var countedDates = BusinessDayDatesBeforeAward(samDate, awardDate);
                count = countedDates.Count;
                isAfterAward = samDate.Date > awardDate.Date;
                countedDatesText = FormatDateList(countedDates);
            }
            else if (TryExtractSamCountAndWindow(evidence, out count, out requiredWindow))
            {
                isAfterAward = false;
                countedDatesText = "";
                samFile = "SAAD";
                awardFile = "SF1449 award";
                issuedByCode = "";
            }
            else
            {
                return;
            }

            var passes = !isAfterAward && count <= requiredWindow;
            var reason = isAfterAward
                ? "the SAM exclusion search date is after the SF1449 award/effective date"
                : $"the computed business-day count is {count}, which is outside the required {requiredWindow}-business-day window";
            existing["Check"] = "SAM Exclusion Search Date";
            existing["Result"] = passes ? "Pass" : "Flag";
            existing["Agent Recommendation"] = passes ? "No action needed" : "Review";
            existing["Fields Reviewed"] = "SAM exclusion search date, award date, issued-by code";
            existing["Recommended Action"] = passes ? "No action needed" : "Review required";
            existing["Data Completeness Notes"] = passes
                ? ""
                : reason + ".";
            existing["Result Summary"] = passes
                ? $"The SAM exclusion search date check passes because the computed business-day count is {count}, which is within the required {requiredWindow}-business-day window."
                : $"The SAM exclusion search date check is flagged because {reason}.";
            if (hasDates)
            {
                existing["Values Compared"] =
                    $"SAAD {samFile} date {FormatDate(samDate)}; SF1449 award {awardFile} date {FormatDate(awardDate)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count {count}; counted business days {countedDatesText}";
                existing["Source Documents"] = JoinDistinct(new[] { samFile, awardFile });
            }
        }

        private static bool TryExtractSamDatesAndWindow(
            string evidence,
            out DateTime samDate,
            out DateTime awardDate,
            out int requiredWindow,
            out string issuedByCode,
            out string samFile,
            out string awardFile)
        {
            samDate = default;
            awardDate = default;
            requiredWindow = 0;
            issuedByCode = "";
            samFile = "SAAD";
            awardFile = "SF1449 award";
            var text = evidence ?? "";
            var samMatch = Regex.Match(
                text,
                @"SAAD\s+(?<file>[^.;\r\n]+?)\s+date\s+(?<date>\d{1,2}/\d{1,2}/\d{2,4})",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var awardMatch = Regex.Match(
                text,
                @"SF1449\s+award\s+(?<file>[^.;\r\n]+?)\s+date\s+(?<date>\d{1,2}/\d{1,2}/\d{2,4})",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var windowMatch = Regex.Match(
                text,
                @"(?:required\s+window|requires\s+a|required)\s*(\d+)\s*-?\s*business",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var issuedByMatch = Regex.Match(
                text,
                @"issued-by\s+code\s+(?<code>[A-Z0-9]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (!samMatch.Success ||
                !awardMatch.Success ||
                !windowMatch.Success ||
                !TryParseDate(samMatch.Groups["date"].Value, out samDate) ||
                !TryParseDate(awardMatch.Groups["date"].Value, out awardDate) ||
                !int.TryParse(windowMatch.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out requiredWindow))
            {
                return false;
            }

            samFile = samMatch.Groups["file"].Value.Trim();
            awardFile = awardMatch.Groups["file"].Value.Trim();
            issuedByCode = issuedByMatch.Success ? issuedByMatch.Groups["code"].Value.Trim().ToUpperInvariant() : "";
            return true;
        }

        private static void ApplySamResultConsistencyGuard(List<Dictionary<string, string>> rows)
        {
            var existing = rows.FirstOrDefault(row =>
                MapCustomerCheck(FirstNonEmpty(Get(row, "Check"), Get(row, "CheckName"), Get(row, "Review Item")))
                    .Equals("SAM Exclusion Search Date", StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                return;
            }

            var evidence = string.Join(" ",
                Get(existing, "Values Compared"),
                Get(existing, "Result Summary"),
                Get(existing, "Data Completeness Notes"));

            if (!TryExtractSamCountAndWindow(evidence, out var count, out var requiredWindow))
            {
                return;
            }

            var passes = count <= requiredWindow;
            existing["Check"] = "SAM Exclusion Search Date";
            existing["Result"] = passes ? "Pass" : "Flag";
            existing["Agent Recommendation"] = passes ? "No action needed" : "Review";
            existing["Fields Reviewed"] = "SAM exclusion search date, award date, issued-by code";
            existing["Recommended Action"] = passes ? "No action needed" : "Review required";

            if (passes)
            {
                existing["Data Completeness Notes"] = "";
                existing["Result Summary"] = $"The SAM exclusion search date check passes because the computed business-day count is {count}, which is within the required {requiredWindow}-business-day window.";
            }
            else
            {
                var reason = $"the computed business-day count is {count}, which is outside the required {requiredWindow}-business-day window";
                existing["Data Completeness Notes"] = reason + ".";
                existing["Result Summary"] = $"The SAM exclusion search date check is flagged because {reason}.";
            }
        }

        private static bool TryExtractSamCountAndWindow(string evidence, out int count, out int requiredWindow)
        {
            count = 0;
            requiredWindow = 0;
            var text = evidence ?? "";
            var countMatch = Regex.Match(
                text,
                @"computed(?:\s+business-day)?(?:\s+count|\s+conclusion)?(?:\s+is|:)?\s*(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var windowMatch = Regex.Match(
                text,
                @"(?:required\s+window|requires\s+a|required)\s*(\d+)\s*-?\s*business",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            return countMatch.Success &&
                   windowMatch.Success &&
                   int.TryParse(countMatch.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out count) &&
                   int.TryParse(windowMatch.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out requiredWindow);
        }

        private static bool TryBuildSamValidationRow(string contractDataJson, out Dictionary<string, string> row)
        {
            row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!TryReadDocumentArray(contractDataJson, out var documents))
            {
                return false;
            }

            var saad = documents.FirstOrDefault(document =>
                GetJsonString(document, "documentType").Equals("SAAD", StringComparison.OrdinalIgnoreCase));
            if (saad.ValueKind == JsonValueKind.Undefined)
            {
                row = SamRow(
                    "Flag",
                    "Review",
                    "Review required",
                    "SAAD not found",
                    "The SAM exclusion search date check is flagged because no SAAD document record was available in the consolidated DU payload.",
                    "No SAAD document record was available in the consolidated DU payload.",
                    "SAAD missing; SF1449 award not evaluated",
                    "");
                return true;
            }

            var award = documents.FirstOrDefault(document =>
                GetJsonString(document, "documentType").Equals("SF1449", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(GetJsonString(document, "awardEffectiveDate")));

            var samFile = SourceDocumentLabel("SAAD", GetJsonString(saad, "sourceFileName"));
            var samDateText = FirstNonEmpty(
                GetJsonString(saad, "saadSamCheckedDate"),
                GetJsonString(saad, "saadSamCheckedDateRaw"));
            if (award.ValueKind == JsonValueKind.Undefined)
            {
                row = SamRow(
                    "Flag",
                    "Review",
                    "Review required",
                    "award date missing",
                    $"The SAM exclusion search date check is flagged because no SF1449 award record with an award/effective date was available. Values reviewed: SAAD {samFile} date {DisplayDate(samDateText)}; SF1449 award date missing; required window not computed; computed count not computed.",
                    "SF1449 award/effective date was not available in the consolidated DU payload.",
                    $"SAAD {samFile} date {DisplayDate(samDateText)}; SF1449 award date missing; required window not computed; computed count not computed",
                    samFile);
                return true;
            }

            var awardFile = SourceDocumentLabel("SF1449", GetJsonString(award, "sourceFileName"));
            var awardDateText = GetJsonString(award, "awardEffectiveDate");
            var issuedByCode = GetJsonString(award, "issuedByCode").Trim();
            var sourceDocuments = JoinDistinct(new[] { samFile, awardFile });
            var requiredWindow = RequiredSamWindow(issuedByCode);

            if (!TryParseDate(samDateText, out var samDate))
            {
                row = SamRow(
                    "Flag",
                    "Review",
                    "Review required",
                    "SAM date missing or unparseable",
                    $"The SAM exclusion search date check is flagged because the SAAD SAM exclusion search date is missing or cannot be parsed. Values reviewed: SAAD {samFile} date {DisplayDate(samDateText)}; SF1449 award {awardFile} date {DisplayDate(awardDateText)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count not computed.",
                    "SAAD SAM exclusion search date was missing or could not be parsed.",
                    $"SAAD {samFile} date {DisplayDate(samDateText)}; SF1449 award {awardFile} date {DisplayDate(awardDateText)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count not computed",
                    sourceDocuments);
                return true;
            }

            if (!TryParseDate(awardDateText, out var awardDate))
            {
                row = SamRow(
                    "Flag",
                    "Review",
                    "Review required",
                    "award date missing or unparseable",
                    $"The SAM exclusion search date check is flagged because the SF1449 award/effective date is missing or cannot be parsed. Values reviewed: SAAD {samFile} date {FormatDate(samDate)}; SF1449 award {awardFile} date {DisplayDate(awardDateText)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count not computed.",
                    "SF1449 award/effective date was missing or could not be parsed.",
                    $"SAAD {samFile} date {FormatDate(samDate)}; SF1449 award {awardFile} date {DisplayDate(awardDateText)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count not computed",
                    sourceDocuments);
                return true;
            }

            var countedDates = BusinessDayDatesBeforeAward(samDate, awardDate);
            var count = countedDates.Count;
            var countedDatesText = FormatDateList(countedDates);
            var isAfterAward = samDate.Date > awardDate.Date;
            var passes = !isAfterAward && count <= requiredWindow;
            var result = passes ? "Pass" : "Flag";
            var recommendation = passes ? "No action needed" : "Review";
            var action = passes ? "No action needed" : "Review required";
            var reason = isAfterAward
                ? "the SAM exclusion search date is after the SF1449 award/effective date"
                : $"the computed business-day count is {count}, which is outside the {requiredWindow}-business-day window";
            var summary = passes
                ? $"The SAAD SAM exclusion search date is {FormatDate(samDate)}; the SF1449 award date is {FormatDate(awardDate)}; issued-by code {DisplayValue(issuedByCode)} requires a {requiredWindow}-business-day window; the computed business-day count is {count}, so the check passes."
                : $"The SAM exclusion search date check is flagged because {reason}. Values reviewed: SAAD {samFile} date {FormatDate(samDate)}; SF1449 award {awardFile} date {FormatDate(awardDate)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count {count}.";
            var dataNotes = passes ? "" : reason + ".";
            var valuesCompared = $"SAAD {samFile} date {FormatDate(samDate)}; SF1449 award {awardFile} date {FormatDate(awardDate)}; issued-by code {DisplayValue(issuedByCode)}; required window {requiredWindow} business days; computed count {count}; counted business days {countedDatesText}";

            row = SamRow(result, recommendation, action, reason, summary, dataNotes, valuesCompared, sourceDocuments);
            return true;
        }

        private static Dictionary<string, string> SamRow(
            string result,
            string recommendation,
            string action,
            string reason,
            string summary,
            string dataCompletenessNotes,
            string valuesCompared,
            string sourceDocuments) =>
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Check"] = "SAM Exclusion Search Date",
                ["Result"] = result,
                ["Agent Recommendation"] = recommendation,
                ["Fields Reviewed"] = "SAM exclusion search date, award date, issued-by code",
                ["Values Compared"] = valuesCompared,
                ["Result Summary"] = summary,
                ["Recommended Action"] = action,
                ["Data Completeness Notes"] = dataCompletenessNotes,
                ["Source Documents"] = sourceDocuments
            };

        private static bool TryReadDocumentArray(string contractDataJson, out List<JsonElement> documents)
        {
            documents = new List<JsonElement>();
            var payload = NormalizeContractPayload(contractDataJson);
            if (string.IsNullOrWhiteSpace(payload))
            {
                return false;
            }

            try
            {
                using var document = JsonDocument.Parse(payload);
                if (!TryGetProperty(document.RootElement, "documents", out var documentArray) ||
                    documentArray.ValueKind != JsonValueKind.Array)
                {
                    return false;
                }

                documents = documentArray.EnumerateArray()
                    .Where(item => item.ValueKind == JsonValueKind.Object)
                    .Select(item => item.Clone())
                    .ToList();
                return documents.Count > 0;
            }
            catch (JsonException)
            {
                return false;
            }
        }

        private static string NormalizeContractPayload(string contractDataJson)
        {
            var payload = (contractDataJson ?? "").Trim();
            const string prefix = "DLA_JSON_PAYLOAD:";
            if (payload.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                payload = payload.Substring(prefix.Length).Trim();
            }

            if (!payload.StartsWith("{", StringComparison.Ordinal))
            {
                return payload;
            }

            try
            {
                using var document = JsonDocument.Parse(payload);
                if (TryGetProperty(document.RootElement, "payload", out var payloadElement))
                {
                    return NormalizeContractPayload(ElementToString(payloadElement));
                }
            }
            catch (JsonException)
            {
                return payload;
            }

            return payload;
        }

        private static int RequiredSamWindow(string issuedByCode)
        {
            var code = (issuedByCode ?? "").Trim().ToUpperInvariant();
            return code is "SPRPA1" or "SPRMM1" or "SPRDL1" or "SPRBL1" or "SPMYM1"
                ? 7
                : 4;
        }

        private static void ApplySemanticAlignmentConsistencyGuard(List<Dictionary<string, string>> rows, string contractDataJson)
        {
            var existing = rows.FirstOrDefault(row =>
                MapCustomerCheck(FirstNonEmpty(Get(row, "Check"), Get(row, "CheckName"), Get(row, "Review Item")))
                    .Equals("Semantic Alignment", StringComparison.OrdinalIgnoreCase));

            if (!TryBuildSemanticAlignmentCorrection(contractDataJson, out var correction) &&
                (existing == null || !TryBuildSemanticAlignmentCorrection(existing, out correction)))
            {
                return;
            }

            if (existing == null)
            {
                rows.Add(correction);
                return;
            }

            existing["Check"] = "Semantic Alignment";
            foreach (var pair in correction)
            {
                existing[pair.Key] = pair.Value;
            }
        }

        private static bool TryBuildSemanticAlignmentCorrection(string contractDataJson, out Dictionary<string, string> row)
        {
            row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!TryReadDocumentArray(contractDataJson, out var documents))
            {
                return false;
            }

            var dd2579 = documents.FirstOrDefault(document =>
                GetJsonString(document, "documentType").Equals("DD2579", StringComparison.OrdinalIgnoreCase));
            if (dd2579.ValueKind == JsonValueKind.Undefined)
            {
                return false;
            }

            var dd2579File = SourceDocumentLabel("DD2579", GetJsonString(dd2579, "sourceFileName"));
            var description = GetJsonString(dd2579, "itemServiceDescription");
            var psc = GetJsonString(dd2579, "productOrServiceCode").Trim().ToUpperInvariant();
            var naics = GetJsonString(dd2579, "naicsCode").Trim();
            if (!PscJ011ConflictsWithAircraftRepair(psc, description))
            {
                return false;
            }

            row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Check"] = "Semantic Alignment",
                ["Result"] = "Flag",
                ["Agent Recommendation"] = "Review",
                ["Fields Reviewed"] = "Item/service description, PSC, NAICS",
                ["Values Compared"] =
                    $"{dd2579File} item/service description: {DisplayValue(description)}; {dd2579File} PSC: {DisplayValue(psc)}; {dd2579File} NAICS: {DisplayValue(naics)}; PSC Manual evidence: J011 is maintenance, repair, and rebuilding of equipment - nuclear ordnance; expected aircraft component/accessory maintenance and repair PSC family is J016. NAICS Manual evidence: {DisplayValue(naics)} should still be evaluated against the item/service description.",
                ["Result Summary"] =
                    $"The check is flagged because the DD2579 item/service description describes aircraft component inspection and repair services, but PSC {psc} maps to maintenance, repair, and rebuilding of nuclear ordnance rather than aircraft components and accessories.",
                ["Recommended Action"] =
                    "Review DD2579 Block 7b and confirm whether the PSC should be J016 for aircraft components and accessories.",
                ["Data Completeness Notes"] =
                    "Semantic Alignment was corrected from the DD2579 evidence and PSC Manual code family because PSC J011 conflicts with aircraft component repair language.",
                ["Source Documents"] = dd2579File
            };
            return true;
        }

        private static bool TryBuildSemanticAlignmentCorrection(
            Dictionary<string, string> existing,
            out Dictionary<string, string> row)
        {
            row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var evidence = string.Join(" ",
                Get(existing, "Values Compared"),
                Get(existing, "Result Summary"),
                Get(existing, "Data Completeness Notes"));

            if (string.IsNullOrWhiteSpace(evidence) ||
                !Regex.IsMatch(evidence, @"\bPSC\s*:\s*J011\b|\bPSC\s+J011\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
                !PscJ011ConflictsWithAircraftRepair("J011", evidence))
            {
                return false;
            }

            var sourceDocuments = FirstNonEmpty(Get(existing, "Source Documents"), "DD2579");
            row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Check"] = "Semantic Alignment",
                ["Result"] = "Flag",
                ["Agent Recommendation"] = "Review",
                ["Fields Reviewed"] = "Item/service description, PSC, NAICS",
                ["Values Compared"] =
                    $"{FirstNonEmpty(Get(existing, "Values Compared"), "DD2579 item/service description references aircraft component inspection and repair; DD2579 PSC J011; DD2579 NAICS 488190")}; PSC Manual evidence: J011 is maintenance, repair, and rebuilding of equipment - nuclear ordnance; expected aircraft component/accessory maintenance and repair PSC family is J016.",
                ["Result Summary"] =
                    "The check is flagged because the DD2579 item/service description describes aircraft component inspection and repair services, but PSC J011 maps to maintenance, repair, and rebuilding of nuclear ordnance rather than aircraft components and accessories.",
                ["Recommended Action"] =
                    "Review DD2579 Block 7b and confirm whether the PSC should be J016 for aircraft components and accessories.",
                ["Data Completeness Notes"] =
                    "Semantic Alignment was corrected from the DD2579 evidence and PSC Manual code family because PSC J011 conflicts with aircraft component repair language.",
                ["Source Documents"] = sourceDocuments
            };
            return true;
        }

        private static bool PscJ011ConflictsWithAircraftRepair(string psc, string description)
        {
            if (!psc.Equals("J011", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(description))
            {
                return false;
            }

            var normalizedDescription = description.ToUpperInvariant();
            var describesAircraft = normalizedDescription.Contains("AIRCRAFT", StringComparison.Ordinal) ||
                                    normalizedDescription.Contains("TURBOPROP", StringComparison.Ordinal) ||
                                    normalizedDescription.Contains("ENGINE", StringComparison.Ordinal);
            var describesRepair = normalizedDescription.Contains("COMPONENT", StringComparison.Ordinal) ||
                                  normalizedDescription.Contains("INSPECTION", StringComparison.Ordinal) ||
                                  normalizedDescription.Contains("REPAIR", StringComparison.Ordinal) ||
                                  normalizedDescription.Contains("MAINTENANCE", StringComparison.Ordinal) ||
                                  normalizedDescription.Contains("REBUILD", StringComparison.Ordinal);

            return describesAircraft && describesRepair;
        }

        private static int CountBusinessDaysBeforeAward(DateTime samDate, DateTime awardDate) =>
            BusinessDayDatesBeforeAward(samDate, awardDate).Count;

        private static List<DateTime> BusinessDayDatesBeforeAward(DateTime samDate, DateTime awardDate)
        {
            var dates = new List<DateTime>();
            if (samDate.Date > awardDate.Date)
            {
                return dates;
            }

            for (var date = samDate.Date; date < awardDate.Date; date = date.AddDays(1))
            {
                if (date.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday)
                {
                    dates.Add(date);
                }
            }

            return dates;
        }

        private static bool TryParseDate(string value, out DateTime parsed)
        {
            parsed = default;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var text = value.Trim();
            var formats = new[]
            {
                "M/d/yyyy",
                "MM/dd/yyyy",
                "M/d/yyyy HH:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "M/d/yyyy h:mm:ss tt",
                "MM/dd/yyyy h:mm:ss tt",
                "yyyy-MM-dd",
                "yyyy-MM-ddTHH:mm:ss",
                "yyyy-MM-ddTHH:mm:ss.fff",
                "dd MMM yyyy",
                "dd MMMM yyyy",
                "MMM d yyyy",
                "MMMM d yyyy"
            };

            return DateTime.TryParseExact(
                    text,
                    formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out parsed)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out parsed);
        }

        private static string FormatDate(DateTime value) =>
            value.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

        private static string FormatDateList(IEnumerable<DateTime> values)
        {
            var dates = values.Select(FormatDate).ToList();
            return dates.Count == 0 ? "none" : string.Join(", ", dates);
        }

        private static string DisplayDate(string value) =>
            string.IsNullOrWhiteSpace(value) ? "missing" : value.Trim();

        private static string DisplayValue(string value) =>
            string.IsNullOrWhiteSpace(value) ? "missing" : value.Trim();

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

        private static string SerializeRows(List<Dictionary<string, string>> rows)
        {
            var normalizedRows = rows.Select(row =>
            {
                var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var column in CustomerColumns)
                {
                    normalized[column] = Get(row, column);
                }

                return normalized;
            });

            return JsonSerializer.Serialize(normalizedRows);
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

            var sourceDocuments = SourceDocumentRows(findings, agentContentJson);
            var sourceDocumentSummary = SourceDocumentSummary(sourceDocuments);
            var reportRows = BuildCustomerRows(findings, sourceDocumentSummary);
            var customerGrid = CustomerSummaryGrid(summary, reportRows, sourceDocuments.Count);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
            AddText(archive, "[Content_Types].xml", ContentTypesXml());
            AddText(archive, "_rels/.rels", RootRelationshipsXml());
            AddText(archive, "xl/workbook.xml", WorkbookXml());
            AddText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml());
            AddText(archive, "xl/styles.xml", StylesXml());
            AddText(archive, "xl/worksheets/sheet1.xml", SheetGridXml("Contract Quality Review Results", customerGrid));
        }

        private static IReadOnlyList<IReadOnlyList<string>> CustomerSummaryGrid(
            IReadOnlyDictionary<string, string> summary,
            IReadOnlyList<Dictionary<string, string>> findings,
            int sourceDocumentCount)
        {
            var passCount = findings.Count(r => Get(r, "Result").Equals("Pass", StringComparison.OrdinalIgnoreCase));
            var flagCount = findings.Count(r => Get(r, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase));
            var notApplicableCount = findings.Count(r => Get(r, "Result").Equals("Not Applicable", StringComparison.OrdinalIgnoreCase));
            var overallStatus = CustomerOverallStatus(findings);
            var contractPackage = FirstNonEmpty(Get(summary, "contract_package"), Get(summary, "contract_batch_id"), Get(summary, "contractBatchId"), "Not specified");
            var created = FirstNonEmpty(Get(summary, "created"), DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'"));
            var documentsProcessed = FirstNonEmpty(Get(summary, "documents_processed"), sourceDocumentCount > 0 ? sourceDocumentCount.ToString() : "");
            var fieldsReviewed = FirstNonEmpty(Get(summary, "fields_reviewed"), $"{CountReviewedFields(findings)} field groups");
            var rulesEvaluated = FirstNonEmpty(Get(summary, "business_rules_evaluated"), findings.Count.ToString());
            var resultSummary = $"{passCount} Pass | {flagCount} Flag | {notApplicableCount} Not Applicable";

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

            return rows;
        }

        private static List<Dictionary<string, string>> BuildCustomerRows(IReadOnlyList<Dictionary<string, string>> findings, string sourceDocumentSummary)
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
                    ? MissingCheckRow(requiredCheck, sourceDocumentSummary)
                    : MergeCustomerRows(requiredCheck, matchingRows, sourceDocumentSummary));
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
                ["Result Summary"] = NormalizeIssue(result, FirstNonEmpty(Get(row, "Result Summary"), issue)),
                ["Recommended Action"] = NormalizeAction(result, action, check),
                ["Data Completeness Notes"] = FirstNonEmpty(Get(row, "Data Completeness Notes"), Get(row, "Data Completeness Note")),
                ["Source Documents"] = FirstNonEmpty(Get(row, "Source Documents"), Get(row, "Documents Reviewed"), Get(row, "SourceFileName"))
            };
        }

        private static Dictionary<string, string> MergeCustomerRows(string check, IReadOnlyList<Dictionary<string, string>> rows, string sourceDocumentSummary)
        {
            var result = rows.Any(r => Get(r, "Result").Equals("Flag", StringComparison.OrdinalIgnoreCase))
                ? "Flag"
                : rows.Any(r => Get(r, "Result").Equals("Pass", StringComparison.OrdinalIgnoreCase))
                    ? "Pass"
                    : "Not Applicable";

            var issue = JoinDistinct(rows.Select(r => FirstNonEmpty(Get(r, "Result Summary"), Get(r, "Issue"))));
            if (string.IsNullOrWhiteSpace(issue) && result.Equals("Pass", StringComparison.OrdinalIgnoreCase))
            {
                issue = "No issue identified";
            }
            var normalizedIssue = CleanCustomerText(NormalizeIssue(result, issue));

            return new Dictionary<string, string>
            {
                ["Check"] = check,
                ["Result"] = result,
                ["Agent Recommendation"] = NormalizeRecommendation(JoinDistinct(rows.Select(r => Get(r, "Agent Recommendation"))), result, issue),
                ["Fields Reviewed"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Fields Reviewed"))), DefaultFieldsReviewed(check))),
                ["Values Compared"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Values Compared"))), DefaultValuesCompared(check))),
                ["Result Summary"] = normalizedIssue,
                ["Recommended Action"] = CleanCustomerText(NormalizeAction(result, JoinDistinct(rows.Select(r => Get(r, "Recommended Action"))), check)),
                ["Data Completeness Notes"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Data Completeness Notes"))), DataCompletenessNote(result, normalizedIssue, JoinDistinct(rows.Select(r => Get(r, "Agent Recommendation")))))),
                ["Source Documents"] = CleanCustomerText(FirstNonEmpty(JoinDistinct(rows.Select(r => Get(r, "Source Documents"))), sourceDocumentSummary))
            };
        }

        private static Dictionary<string, string> MissingCheckRow(string check, string sourceDocumentSummary) =>
            new()
            {
                ["Check"] = check,
                ["Result"] = "Flag",
                ["Agent Recommendation"] = "Missing evidence",
                ["Fields Reviewed"] = DefaultFieldsReviewed(check),
                ["Values Compared"] = DefaultValuesCompared(check),
                ["Result Summary"] = "The agent did not return enough evidence to complete this check.",
                ["Recommended Action"] = DefaultRecommendedAction(check),
                ["Data Completeness Notes"] = "The agent did not return a complete row for this required check.",
                ["Source Documents"] = sourceDocumentSummary
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
                        ["Notes"] = FirstNonEmpty(Get(finding, "Result Summary"), Get(finding, "Issue"), Get(finding, "Exception"), Get(finding, "Evidence"))
                    });
                }
            }

            return rows;
        }

        private static string SourceDocumentSummary(IReadOnlyList<Dictionary<string, string>> sourceDocuments)
        {
            return CleanCustomerText(JoinDistinct(sourceDocuments.Select(row =>
            {
                var type = Get(row, "Document Type");
                var file = Get(row, "File Name");
                var notes = Get(row, "Notes");
                var value = SourceDocumentLabel(type, file);
                return string.IsNullOrWhiteSpace(notes) ? value : $"{value} ({notes})";
            })));
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
                return CleanCustomerText(FirstNonEmpty(issue, "No issue identified"));
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
                return "Detailed evidence is available in the source system.";
            }

            var upper = text.ToUpperInvariant();
            if (upper.Contains("STACK TRACE", StringComparison.Ordinal) ||
                upper.Contains("QUEUE ITEM", StringComparison.Ordinal) ||
                upper.Contains("ROBOT ", StringComparison.Ordinal) ||
                upper.Contains("JOB ID", StringComparison.Ordinal))
            {
                return "Detailed processing information is available in the source system.";
            }

            return Truncate(text, CustomerCellLimit);
        }

        private static string JoinDistinct(IEnumerable<string> values) =>
            string.Join("; ", values
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => v.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase));

        private static string DataCompletenessNote(string result, string resultSummary, string recommendation)
        {
            if (!result.Equals("Flag", StringComparison.OrdinalIgnoreCase) ||
                (!recommendation.Equals("Missing evidence", StringComparison.OrdinalIgnoreCase) &&
                 !IndicatesMissingEvidence(resultSummary)))
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

        private static string SourceDocumentLabel(string documentType, string fileName)
        {
            var type = FirstNonEmpty(documentType, "Document");
            var file = Path.GetFileNameWithoutExtension(fileName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(file))
            {
                return type;
            }

            var normalizedType = NormalizeDocumentToken(type);
            var normalizedFile = NormalizeDocumentToken(file);
            return !string.IsNullOrWhiteSpace(normalizedType) && normalizedFile.StartsWith(normalizedType, StringComparison.OrdinalIgnoreCase)
                ? file
                : $"{type}: {file}";
        }

        private static string NormalizeDocumentToken(string value)
        {
            var chars = (value ?? "")
                .Where(char.IsLetterOrDigit)
                .Select(char.ToUpperInvariant)
                .ToArray();
            return new string(chars);
        }

        private static string SheetGridXml(string sheetName, IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<cols>");
            var widths = new[] { 34, 18, 24, 34, 85, 72, 72, 58, 64 };
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
            var isBusinessRuleRow = rowIndex >= 13;
            sb.Append("<row r=\"").Append(rowIndex).Append("\"");
            if (isBusinessRuleRow)
            {
                sb.Append(" ht=\"108\" customHeight=\"1\"");
            }
            sb.Append(">");
            for (var i = 0; i < values.Count; i++)
            {
                var cellRef = ColumnName(i + 1) + rowIndex;
                sb.Append("<c r=\"").Append(cellRef).Append("\" s=\"1\" t=\"inlineStr\"><is><t>");
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
            "</sheets></workbook>";

        private static string WorkbookRelationshipsXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>";

        private static string StylesXml() =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment wrapText=\"1\" vertical=\"top\"/></xf></cellXfs>" +
            "</styleSheet>";
    }
}
