using System;
using System.Globalization;
using System.Text.RegularExpressions;
using UiPath.CodedWorkflows;

namespace IDP_SAAD
{
    public class NormalizeSaadSamDate : CodedWorkflow
    {
        private static readonly Regex DatePattern = new(
            @"\b(?<date>\d{4}(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\d{2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        [Workflow]
        public (string saadSamCheckedDate, string saadSamCheckedDateRaw, string saadSamCheckedDateSource) Execute(
            string in_ExtractedSamDate,
            string in_DocumentText)
        {
            var extracted = NormalizeDate(in_ExtractedSamDate);
            if (!string.IsNullOrWhiteSpace(extracted))
            {
                return (
                    extracted,
                    in_ExtractedSamDate?.Trim() ?? string.Empty,
                    "DU_FIELD");
            }

            var fallbackRaw = FindImmediatelyPriorSamDate(in_DocumentText);
            var fallback = NormalizeDate(fallbackRaw);
            return (
                fallback,
                fallbackRaw,
                string.IsNullOrWhiteSpace(fallback) ? "NOT_FOUND" : "DOCUMENT_TEXT_FALLBACK");
        }

        private static string FindImmediatelyPriorSamDate(string? documentText)
        {
            if (string.IsNullOrWhiteSpace(documentText))
            {
                return string.Empty;
            }

            var text = Regex.Replace(documentText, @"\s+", " ");
            var immediateMarker = Regex.Match(
                text,
                @"Immediately\s+prior\s+to\s+award",
                RegexOptions.IgnoreCase);

            if (immediateMarker.Success)
            {
                var immediateTail = text.Substring(immediateMarker.Index, Math.Min(220, text.Length - immediateMarker.Index));
                var immediateTailDate = DatePattern.Match(immediateTail);
                if (immediateTailDate.Success)
                {
                    return immediateTailDate.Groups["date"].Value;
                }

                var samSectionStart = Regex.Match(
                    text,
                    @"Contracting\s+Officer\s+checked\s+SAM\.gov",
                    RegexOptions.IgnoreCase);
                if (samSectionStart.Success)
                {
                    var samSection = text.Substring(samSectionStart.Index, Math.Min(2400, text.Length - samSectionStart.Index));
                    var sectionDates = DatePattern.Matches(samSection);
                    if (sectionDates.Count >= 2)
                    {
                        return sectionDates[1].Groups["date"].Value;
                    }

                    if (sectionDates.Count == 1)
                    {
                        return sectionDates[0].Groups["date"].Value;
                    }
                }
            }

            var matches = DatePattern.Matches(text);
            return matches.Count >= 2 ? matches[1].Groups["date"].Value : string.Empty;
        }

        private static string NormalizeDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var raw = value.Trim();
            var compact = Regex.Match(raw, @"^(?<year>\d{4})(?<month>[A-Za-z]{3,4})(?<day>\d{2})$");
            if (compact.Success)
            {
                var normalizedMonth = compact.Groups["month"].Value.ToUpperInvariant() == "SEPT"
                    ? "SEP"
                    : compact.Groups["month"].Value;
                raw = $"{compact.Groups["day"].Value} {normalizedMonth} {compact.Groups["year"].Value}";
            }

            var formats = new[]
            {
                "M/d/yyyy",
                "M/d/yy",
                "M-d-yyyy",
                "M-d-yy",
                "yyyy-MM-dd",
                "dd MMM yyyy",
                "dd MMMM yyyy",
                "MMM d yyyy",
                "MMMM d yyyy"
            };

            if (DateTime.TryParseExact(
                    raw,
                    formats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out var parsed)
                || DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out parsed))
            {
                return parsed.ToString("MM/dd/yyyy 00:00:00", CultureInfo.InvariantCulture);
            }

            return string.Empty;
        }
    }
}
