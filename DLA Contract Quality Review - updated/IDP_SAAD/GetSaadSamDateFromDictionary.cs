using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using UiPath.CodedWorkflows;

namespace IDP_SAAD
{
    public class GetSaadSamDateFromDictionary : CodedWorkflow
    {
        [Workflow]
        public (string saadExtractedSamDate, string saadExtractedFieldsJson) Execute(
            object in_DocumentData)
        {
            var data = in_DocumentData?.GetType().GetProperty("Data")?.GetValue(in_DocumentData);
            var fieldSummary = ReadFields(data).ToArray();

            var selected = fieldSummary
                .Where(HasValue)
                .OrderByDescending(IsExactSamCheckedDate)
                .ThenByDescending(IsLikelySamDate)
                .FirstOrDefault(IsPossibleSamDate);

            return (
                selected != null && selected.TryGetValue("Value", out var value) ? ToText(value) : string.Empty,
                JsonConvert.SerializeObject(fieldSummary));
        }

        private static IEnumerable<Dictionary<string, object?>> ReadFields(object? data)
        {
            if (data == null)
            {
                return Array.Empty<Dictionary<string, object?>>();
            }

            var getFields = data.GetType().GetMethod("GetFields", BindingFlags.Instance | BindingFlags.Public, Type.EmptyTypes);
            if (getFields != null)
            {
                var fields = (getFields.Invoke(data, null) as System.Collections.IEnumerable)?.Cast<object>()
                    ?? Array.Empty<object>();

                return fields.Select(field => new Dictionary<string, object?>
                {
                    ["FieldId"] = ReadProperty(field, "FieldId"),
                    ["FieldName"] = ReadProperty(field, "FieldName"),
                    ["Value"] = ReadResultValue(field),
                    ["FieldProperties"] = SnapshotPublicProperties(field),
                    ["FirstValueProperties"] = SnapshotPublicProperties(ReadFirstValueObject(field))
                });
            }

            return data.GetType()
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.GetIndexParameters().Length == 0)
                .Select(property =>
                {
                    var field = SafeGet(property, data);
                    return new Dictionary<string, object?>
                    {
                        ["FieldId"] = property.Name,
                        ["FieldName"] = property.Name,
                        ["Value"] = ReadGeneratedFieldValue(field),
                        ["FieldProperties"] = SnapshotPublicProperties(field),
                        ["FirstValueProperties"] = new Dictionary<string, object?>()
                    };
                });
        }

        private static bool HasValue(Dictionary<string, object?> field)
        {
            return !string.IsNullOrWhiteSpace(ToText(field.GetValueOrDefault("Value")));
        }

        private static bool IsPossibleSamDate(Dictionary<string, object?> field)
        {
            var label = Label(field);
            return label.Contains("sam", StringComparison.OrdinalIgnoreCase)
                || label.Contains("registration", StringComparison.OrdinalIgnoreCase)
                || label.Contains("checked", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsExactSamCheckedDate(Dictionary<string, object?> field)
        {
            var label = Label(field);
            return label.Equals("HasAnActiveRegistrationInSAMDateSAMChecked", StringComparison.OrdinalIgnoreCase)
                || label.Contains("DateSAMChecked", StringComparison.OrdinalIgnoreCase)
                || (label.Contains("sam", StringComparison.OrdinalIgnoreCase)
                    && label.Contains("checked", StringComparison.OrdinalIgnoreCase)
                    && label.Contains("date", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsLikelySamDate(Dictionary<string, object?> field)
        {
            var label = Label(field);
            return label.Contains("sam", StringComparison.OrdinalIgnoreCase)
                && (label.Contains("date", StringComparison.OrdinalIgnoreCase)
                    || label.Contains("checked", StringComparison.OrdinalIgnoreCase));
        }

        private static string Label(Dictionary<string, object?> field)
        {
            return $"{ToText(field.GetValueOrDefault("FieldId"))} {ToText(field.GetValueOrDefault("FieldName"))}";
        }

        private static object? ReadGeneratedFieldValue(object? field)
        {
            return ReadProperty(field, "Value");
        }

        private static object? ReadResultValue(object? field)
        {
            var firstValue = ReadFirstValueObject(field);
            return ReadProperty(firstValue, "Value");
        }

        private static object? ReadFirstValueObject(object? field)
        {
            var values = ReadProperty(field, "Values") as System.Collections.IEnumerable;
            return values?.Cast<object>().FirstOrDefault();
        }

        private static object? ReadProperty(object? value, string propertyName)
        {
            return value?.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public)?.GetValue(value);
        }

        private static object? SafeGet(PropertyInfo property, object target)
        {
            try
            {
                return property.GetValue(target);
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<string, object?> SnapshotPublicProperties(object? value)
        {
            if (value == null)
            {
                return new Dictionary<string, object?>();
            }

            var snapshot = new Dictionary<string, object?>();
            foreach (var property in value.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (property.GetIndexParameters().Length > 0)
                {
                    continue;
                }

                object? propertyValue;
                try
                {
                    propertyValue = property.GetValue(value);
                }
                catch
                {
                    continue;
                }

                snapshot[property.Name] = IsSimple(propertyValue)
                    ? propertyValue
                    : propertyValue?.ToString();
            }

            return snapshot;
        }

        private static bool IsSimple(object? value)
        {
            if (value == null)
            {
                return true;
            }

            var type = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(Guid);
        }

        private static string ToText(object? value)
        {
            return value switch
            {
                null => string.Empty,
                DateTime date => date.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture),
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => value.ToString() ?? string.Empty
            };
        }
    }
}
