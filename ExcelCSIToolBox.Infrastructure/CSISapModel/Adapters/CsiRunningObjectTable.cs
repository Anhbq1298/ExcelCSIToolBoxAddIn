using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.Adapters
{
    internal static class CsiRunningObjectTable
    {
        private const int S_OK = 0;

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable runningObjectTable);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);

        internal static IReadOnlyList<CsiAttachResult> GetRunningInstances<TCsiObject, TSapModel>(
            string productName,
            string progId,
            Func<TCsiObject, TSapModel> getSapModel)
            where TCsiObject : class
            where TSapModel : class
        {
            var instances = new List<CsiAttachResult>();
            IRunningObjectTable runningObjectTable;
            if (GetRunningObjectTable(0, out runningObjectTable) != S_OK || runningObjectTable == null)
            {
                return instances;
            }

            IEnumMoniker enumMoniker = null;
            try
            {
                runningObjectTable.EnumRunning(out enumMoniker);
                if (enumMoniker == null)
                {
                    return instances;
                }

                enumMoniker.Reset();
                var monikers = new IMoniker[1];
                while (enumMoniker.Next(1, monikers, IntPtr.Zero) == S_OK)
                {
                    IMoniker moniker = monikers[0];
                    if (moniker == null)
                    {
                        continue;
                    }

                    string displayName = GetDisplayName(moniker);
                    if (!IsCandidate(displayName, productName, progId))
                    {
                        continue;
                    }

                    object runningObject;
                    try
                    {
                        runningObjectTable.GetObject(moniker, out runningObject);
                    }
                    catch
                    {
                        continue;
                    }

                    TCsiObject csiObject = runningObject as TCsiObject;
                    if (csiObject == null)
                    {
                        ReleaseComReference(runningObject);
                        continue;
                    }

                    TSapModel sapModel = null;
                    try
                    {
                        sapModel = getSapModel(csiObject);
                    }
                    catch
                    {
                    }

                    if (sapModel == null)
                    {
                        ReleaseComReference(csiObject);
                        continue;
                    }

                    instances.Add(CsiAttachResult.Success(csiObject, sapModel, null, displayName, ExtractProcessId(displayName)));
                }
            }
            finally
            {
                ReleaseComReference(enumMoniker);
                ReleaseComReference(runningObjectTable);
            }

            return instances;
        }

        private static string GetDisplayName(IMoniker moniker)
        {
            IBindCtx bindContext = null;
            try
            {
                if (CreateBindCtx(0, out bindContext) != S_OK || bindContext == null)
                {
                    return string.Empty;
                }

                string displayName;
                moniker.GetDisplayName(bindContext, null, out displayName);
                return displayName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComReference(bindContext);
            }
        }

        private static bool IsCandidate(string displayName, string productName, string progId)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            return displayName.IndexOf(progId, StringComparison.OrdinalIgnoreCase) >= 0
                || displayName.IndexOf(productName, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int? ExtractProcessId(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return null;
            }

            Match match = Regex.Match(displayName, @"(?:pid|process\s*id|processid)\D*(\d+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(displayName, @"\b(\d{3,})\b");
            }

            int processId;
            return match.Success && int.TryParse(match.Groups[1].Value, out processId)
                ? (int?)processId
                : null;
        }

        private static void ReleaseComReference(object comReference)
        {
            if (comReference == null || !Marshal.IsComObject(comReference))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comReference);
            }
            catch
            {
            }
        }
    }
}
