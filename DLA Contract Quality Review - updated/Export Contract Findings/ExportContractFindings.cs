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

            WriteWorkbook(localWorkbookPath, rows, content);

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
            IReadOnlyDictionary<string, string> summary)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            var reportRows = findings.Select(ToReportRow).ToList();
            var sourceDocuments = SourceDocumentRows(findings);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
            AddText(archive, "[Content_Types].xml", ContentTypesXml());
            AddText(archive, "_rels/.rels", RootRelationshipsXml());
            AddText(archive, "xl/workbook.xml", WorkbookXml());
            AddText(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml());
            AddText(archive, "xl/styles.xml", StylesXml());
            AddText(archive, "xl/worksheets/sheet1.xml", SheetXml("Review Summary", SummaryRows(summary, reportRows)));
            AddText(archive, "xl/worksheets/sheet2.xml", SheetXml("Findings", reportRows));
            AddText(archive, "xl/worksheets/sheet3.xml", SheetXml("Source Documents", sourceDocuments));
        }

        private static IReadOnlyList<Dictionary<string, string>> SummaryRows(
            IReadOnlyDictionary<string, string> summary,
            IReadOnlyList<Dictionary<string, string>> findings)
        {
            var rows = new List<Dictionary<string, string>>
            {
                new() { ["Item"] = "Report Created", ["Details"] = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss 'UTC'") },
                new() { ["Item"] = "Overall Result", ["Details"] = FirstNonEmpty(Get(summary, "overall_result"), HighestResult(findings)) },
                new() { ["Item"] = "Total Findings", ["Details"] = findings.Count.ToString() },
                new() { ["Item"] = "Flagged Items", ["Details"] = findings.Count(r => Get(r, "Result").Equals("FLAG", StringComparison.OrdinalIgnoreCase)).ToString() },
                new() { ["Item"] = "Items Needing Review", ["Details"] = findings.Count(r => Get(r, "Result").Equals("REVIEW", StringComparison.OrdinalIgnoreCase)).ToString() },
                new() { ["Item"] = "Data Source", ["Details"] = FriendlyDataSource(FirstNonEmpty(Get(summary, "data_source_mode"), "IDP_JSON_PRIMARY")) },
                new() { ["Item"] = "Data Fabric Used", ["Details"] = FriendlyBoolean(Get(summary, "data_fabric_source_used")) },
                new() { ["Item"] = "Review Batch", ["Details"] = Get(summary, "contract_batch_id") }
            };

            return rows;
        }

        private static Dictionary<string, string> ToReportRow(IReadOnlyDictionary<string, string> row)
        {
            var checkId = FirstNonEmpty(Get(row, "Check"), Get(row, "CheckId"));
            var checkName = FirstNonEmpty(Get(row, "CheckName"), Get(row, "Review Item"));
            var check = string.IsNullOrWhiteSpace(checkId)
                ? checkName
                : string.IsNullOrWhiteSpace(checkName)
                    ? checkId
                    : $"{checkId} - {checkName}";

            return new Dictionary<string, string>
            {
                ["Check"] = check,
                ["Result"] = Get(row, "Result"),
                ["Priority"] = FirstNonEmpty(Get(row, "Priority"), Get(row, "Severity")),
                ["Documents Reviewed"] = FirstNonEmpty(Get(row, "Documents Reviewed"), Get(row, "SourceFileName")),
                ["Field Reviewed"] = FirstNonEmpty(Get(row, "Field Reviewed"), Get(row, "FieldName")),
                ["Values Compared"] = FirstNonEmpty(Get(row, "Values Compared"), Get(row, "ComparedValues"), CombinedValues(row)),
                ["What We Found"] = FirstNonEmpty(Get(row, "What We Found"), Get(row, "Evidence")),
                ["Issue"] = FirstNonEmpty(Get(row, "Issue"), Get(row, "Exception")),
                ["Recommended Action"] = FirstNonEmpty(Get(row, "Recommended Action"), Get(row, "RecommendedAction")),
                ["Data Source"] = FriendlyDataSource(FirstNonEmpty(Get(row, "Data Source"), Get(row, "DataSource"), Get(row, "DataSourceMode"), "IDP_JSON_PRIMARY"))
            };
        }

        private static IReadOnlyList<Dictionary<string, string>> SourceDocumentRows(IReadOnlyList<Dictionary<string, string>> findings)
        {
            var rows = new List<Dictionary<string, string>>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
            "<sheet name=\"Review Summary\" sheetId=\"1\" r:id=\"rId1\"/>" +
            "<sheet name=\"Findings\" sheetId=\"2\" r:id=\"rId2\"/>" +
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
