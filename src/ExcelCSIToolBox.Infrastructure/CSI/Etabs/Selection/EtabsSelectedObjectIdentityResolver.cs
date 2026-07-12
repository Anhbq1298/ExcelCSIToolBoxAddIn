using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ETABSv1;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;
using ExcelCSIToolBox.Core.Contracts.CSI;

namespace ExcelCSIToolBox.Infrastructure.CSI.Etabs.Selection
{
    public sealed class EtabsSelectedObjectIdentityResolver : ISelectedObjectIdentityResolver
    {
        private readonly IEtabsConnectionService _connectionService;

        public EtabsSelectedObjectIdentityResolver(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public OperationResult<IReadOnlyList<CsiObjectIdentity>> ResolveSelectedObjects()
        {
            cSapModel sapModel = _connectionService.SapModel as cSapModel;
            if (sapModel == null)
            {
                return OperationResult<IReadOnlyList<CsiObjectIdentity>>.Failure(
                    "The attached ETABS model is invalid. Please reattach and try again.");
            }

            try
            {
                int numberItems = 0;
                int[] objectTypes = null;
                string[] objectNames = null;
                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret != 0)
                {
                    return OperationResult<IReadOnlyList<CsiObjectIdentity>>.Failure(
                        "Failed to read selected objects from ETABS.");
                }

                var identities = new List<CsiObjectIdentity>();
                if (numberItems <= 0 || objectTypes == null || objectNames == null)
                {
                    return OperationResult<IReadOnlyList<CsiObjectIdentity>>.Success(identities);
                }

                for (int i = 0; i < numberItems; i++)
                {
                    if (i >= objectTypes.Length || i >= objectNames.Length)
                    {
                        continue;
                    }

                    string uniqueName = Clean(objectNames[i]);
                    if (string.IsNullOrWhiteSpace(uniqueName))
                    {
                        continue;
                    }

                    AddSelectedObjectIdentity(sapModel, identities, objectTypes[i], uniqueName);
                }

                return OperationResult<IReadOnlyList<CsiObjectIdentity>>.Success(identities);
            }
            catch (Exception ex)
            {
                return OperationResult<IReadOnlyList<CsiObjectIdentity>>.Failure(
                    "Failed to resolve selected ETABS object identities: " + ex.Message);
            }
        }

        private static void AddSelectedObjectIdentity(
            cSapModel sapModel,
            IList<CsiObjectIdentity> identities,
            int objectType,
            string uniqueName)
        {
            if (objectType == CSISapModelObjectTypeIds.Point)
            {
                string label;
                string story;
                TryGetPointLabel(sapModel, uniqueName, out label, out story);
                AddIdentity(identities, CsiObjectTypes.Point, uniqueName, label, story);
                return;
            }

            if (objectType == CSISapModelObjectTypeIds.Frame)
            {
                string label;
                string story;
                TryGetFrameLabel(sapModel, uniqueName, out label, out story);
                AddIdentity(identities, CsiObjectTypes.Frame, uniqueName, label, story);
                AddPierAndSpandrelIdentitiesForFrame(sapModel, identities, uniqueName, story);
                return;
            }

            if (objectType == CSISapModelObjectTypeIds.Shell)
            {
                string label;
                string story;
                TryGetAreaLabel(sapModel, uniqueName, out label, out story);
                AddIdentity(identities, CsiObjectTypes.Area, uniqueName, label, story);
                AddPierAndSpandrelIdentitiesForArea(sapModel, identities, uniqueName, story);
                return;
            }

            AddIdentity(
                identities,
                CsiObjectTypes.Unknown,
                uniqueName,
                null,
                null,
                new[] { objectType.ToString(CultureInfo.InvariantCulture) });
        }

        private static void AddPierAndSpandrelIdentitiesForFrame(
            cSapModel sapModel,
            IList<CsiObjectIdentity> identities,
            string frameName,
            string story)
        {
            string pierName = string.Empty;
            if (sapModel.FrameObj.GetPier(frameName, ref pierName) == 0 && !string.IsNullOrWhiteSpace(pierName))
            {
                AddIdentity(identities, CsiObjectTypes.Pier, pierName, pierName, story);
            }

            string spandrelName = string.Empty;
            if (sapModel.FrameObj.GetSpandrel(frameName, ref spandrelName) == 0 && !string.IsNullOrWhiteSpace(spandrelName))
            {
                AddIdentity(identities, CsiObjectTypes.Spandrel, spandrelName, spandrelName, story);
            }
        }

        private static void AddPierAndSpandrelIdentitiesForArea(
            cSapModel sapModel,
            IList<CsiObjectIdentity> identities,
            string areaName,
            string story)
        {
            string pierName = string.Empty;
            if (sapModel.AreaObj.GetPier(areaName, ref pierName) == 0 && !string.IsNullOrWhiteSpace(pierName))
            {
                AddIdentity(identities, CsiObjectTypes.Pier, pierName, pierName, story);
            }

            string spandrelName = string.Empty;
            if (sapModel.AreaObj.GetSpandrel(areaName, ref spandrelName) == 0 && !string.IsNullOrWhiteSpace(spandrelName))
            {
                AddIdentity(identities, CsiObjectTypes.Spandrel, spandrelName, spandrelName, story);
            }
        }

        private static void TryGetPointLabel(cSapModel sapModel, string uniqueName, out string label, out string story)
        {
            label = string.Empty;
            story = string.Empty;
            sapModel.PointObj.GetLabelFromName(uniqueName, ref label, ref story);
        }

        private static void TryGetFrameLabel(cSapModel sapModel, string uniqueName, out string label, out string story)
        {
            label = string.Empty;
            story = string.Empty;
            sapModel.FrameObj.GetLabelFromName(uniqueName, ref label, ref story);
        }

        private static void TryGetAreaLabel(cSapModel sapModel, string uniqueName, out string label, out string story)
        {
            label = string.Empty;
            story = string.Empty;
            sapModel.AreaObj.GetLabelFromName(uniqueName, ref label, ref story);
        }

        private static void AddIdentity(
            IList<CsiObjectIdentity> identities,
            string objectType,
            string uniqueName,
            string label,
            string story,
            IEnumerable<string> additionalMatchKeys = null)
        {
            CsiObjectIdentity identity = CsiObjectIdentity.Create(
                objectType,
                uniqueName,
                label,
                story,
                additionalMatchKeys);
            if (identity.MatchKeys == null || identity.MatchKeys.Count == 0)
            {
                return;
            }

            if (identities.Any(existing =>
                existing != null &&
                string.Equals(existing.ObjectType, identity.ObjectType, StringComparison.OrdinalIgnoreCase) &&
                identity.MatchKeys.Any(existing.Matches)))
            {
                return;
            }

            identities.Add(identity);
        }

        private static string Clean(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }
}
