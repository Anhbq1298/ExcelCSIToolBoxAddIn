using System;
using System.Collections.Generic;
using System.IO;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    public sealed class CsiOperationLogger
    {
        private readonly string _logFilePath;

        public CsiOperationLogger()
        {
            string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string folder = Path.Combine(root, "ExcelCSIToolBoxAddIn");
            _logFilePath = Path.Combine(folder, "mcp-write-operations.log");
        }

        public void Log(
            string productType,
            string operationName,
            string category,
            string subCategory,
            CsiMethodRiskLevel riskLevel,
            string argumentsSummary,
            IReadOnlyList<string> affectedObjects,
            bool confirmed,
            bool succeeded,
            string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));

                string affected = affectedObjects == null
                    ? string.Empty
                    : string.Join(",", affectedObjects);

                string line =
                    DateTimeOffset.Now.ToString("o") + "\t" +
                    Safe(productType) + "\t" +
                    Safe(operationName) + "\t" +
                    Safe(category) + "\t" +
                    Safe(subCategory) + "\t" +
                    riskLevel + "\t" +
                    Safe(argumentsSummary) + "\t" +
                    Safe(affected) + "\t" +
                    confirmed + "\t" +
                    succeeded + "\t" +
                    Safe(message);

                File.AppendAllText(_logFilePath, line + Environment.NewLine);
            }
            catch
            {
                // Logging must never block or crash Excel.
            }
        }

        private static string Safe(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ");
        }
    }
}
