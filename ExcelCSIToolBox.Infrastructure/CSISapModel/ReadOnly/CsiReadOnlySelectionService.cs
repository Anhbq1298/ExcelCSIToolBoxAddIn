using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.ReadOnly
{
    /// <summary>
    /// Read-only selection service that reads selected objects and frame names from
    /// a running CSI model without modifying it.
    ///
    /// The raw SapModel COM objects stay entirely inside this class and are never exposed
    /// to AI/MCP/UI layers.
    /// </summary>
    public class CsiReadOnlySelectionService : ICsiReadOnlySelectionService
    {
        private readonly ICSISapModelConnectionAdapter<ETABSv1.cSapModel>   _etabsAdapter;
        private readonly ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> _sap2000Adapter;

        public CsiReadOnlySelectionService()
        {
            _etabsAdapter   = CSISapModelConnectionAdapterFactory.CreateEtabs();
            _sap2000Adapter = CSISapModelConnectionAdapterFactory.CreateSap2000();
        }

        /// <summary>
        /// Return all selected objects from whichever model is currently attached.
        /// Tries ETABS first, then SAP2000. Does not modify the model.
        /// </summary>
        public OperationResult<List<CsiSelectedObjectDto>> GetSelectedObjects()
        {
            // Try ETABS.
            OperationResult<ETABSv1.cSapModel> etabsModel = _etabsAdapter.EnsureSapModel();
            if (etabsModel.IsSuccess)
            {
                return GetSelectedObjectsFromEtabs(etabsModel.Data);
            }

            // Try SAP2000.
            OperationResult<SAP2000v1.cSapModel> sap2000Model = _sap2000Adapter.EnsureSapModel();
            if (sap2000Model.IsSuccess)
            {
                return GetSelectedObjectsFromSap2000(sap2000Model.Data);
            }

            return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                "No running ETABS or SAP2000 instance found. Please open a model first.");
        }

        /// <summary>
        /// Return unique names of currently selected frame objects only.
        /// Does not modify the model.
        /// </summary>
        public OperationResult<List<string>> GetSelectedFrameNames()
        {
            // Try ETABS.
            OperationResult<ETABSv1.cSapModel> etabsModel = _etabsAdapter.EnsureSapModel();
            if (etabsModel.IsSuccess)
            {
                return GetSelectedFramesFromEtabs(etabsModel.Data);
            }

            // Try SAP2000.
            OperationResult<SAP2000v1.cSapModel> sap2000Model = _sap2000Adapter.EnsureSapModel();
            if (sap2000Model.IsSuccess)
            {
                return GetSelectedFramesFromSap2000(sap2000Model.Data);
            }

            return OperationResult<List<string>>.Failure(
                "No running ETABS or SAP2000 instance found. Please open a model first.");
        }

        // ─── ETABS helpers ───────────────────────────────────────────────────────

        private static OperationResult<List<CsiSelectedObjectDto>> GetSelectedObjectsFromEtabs(
            ETABSv1.cSapModel sapModel)
        {
            try
            {
                int numberItems      = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret != 0)
                {
                    return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                        "ETABS SelectObj.GetSelected returned an error.");
                }

                return BuildSelectedObjectList(numberItems, objectTypes, objectNames);
            }
            catch (Exception ex)
            {
                return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                    "Failed to read selected objects from ETABS: " + ex.Message);
            }
        }

        private static OperationResult<List<string>> GetSelectedFramesFromEtabs(
            ETABSv1.cSapModel sapModel)
        {
            try
            {
                int numberItems      = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret != 0)
                {
                    return OperationResult<List<string>>.Failure(
                        "ETABS SelectObj.GetSelected returned an error.");
                }

                return BuildSelectedFrameList(numberItems, objectTypes, objectNames);
            }
            catch (Exception ex)
            {
                return OperationResult<List<string>>.Failure(
                    "Failed to read selected frames from ETABS: " + ex.Message);
            }
        }

        // ─── SAP2000 helpers ─────────────────────────────────────────────────────

        private static OperationResult<List<CsiSelectedObjectDto>> GetSelectedObjectsFromSap2000(
            SAP2000v1.cSapModel sapModel)
        {
            try
            {
                int numberItems      = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret != 0)
                {
                    return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                        "SAP2000 SelectObj.GetSelected returned an error.");
                }

                return BuildSelectedObjectList(numberItems, objectTypes, objectNames);
            }
            catch (Exception ex)
            {
                return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                    "Failed to read selected objects from SAP2000: " + ex.Message);
            }
        }

        private static OperationResult<List<string>> GetSelectedFramesFromSap2000(
            SAP2000v1.cSapModel sapModel)
        {
            try
            {
                int numberItems      = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (ret != 0)
                {
                    return OperationResult<List<string>>.Failure(
                        "SAP2000 SelectObj.GetSelected returned an error.");
                }

                return BuildSelectedFrameList(numberItems, objectTypes, objectNames);
            }
            catch (Exception ex)
            {
                return OperationResult<List<string>>.Failure(
                    "Failed to read selected frames from SAP2000: " + ex.Message);
            }
        }

        // ─── Shared builders ─────────────────────────────────────────────────────

        private static OperationResult<List<CsiSelectedObjectDto>> BuildSelectedObjectList(
            int numberItems,
            int[]    objectTypes,
            string[] objectNames)
        {
            var list = new List<CsiSelectedObjectDto>();

            for (int i = 0; i < numberItems; i++)
            {
                if (objectTypes == null || objectNames == null ||
                    i >= objectTypes.Length || i >= objectNames.Length)
                {
                    continue;
                }

                string typeName = ObjectTypeIdToName(objectTypes[i]);
                list.Add(new CsiSelectedObjectDto
                {
                    ObjectType = typeName,
                    UniqueName = objectNames[i]
                });
            }

            if (list.Count == 0)
            {
                return OperationResult<List<CsiSelectedObjectDto>>.Failure(
                    "No objects are currently selected in the model.");
            }

            return OperationResult<List<CsiSelectedObjectDto>>.Success(list);
        }

        private static OperationResult<List<string>> BuildSelectedFrameList(
            int numberItems,
            int[]    objectTypes,
            string[] objectNames)
        {
            var frameNames = new List<string>();

            for (int i = 0; i < numberItems; i++)
            {
                if (objectTypes == null || objectNames == null ||
                    i >= objectTypes.Length || i >= objectNames.Length)
                {
                    continue;
                }

                // Object type 2 = Frame (CSISapModelObjectTypeIds.Frame).
                if (objectTypes[i] == 2 && !string.IsNullOrWhiteSpace(objectNames[i]))
                {
                    frameNames.Add(objectNames[i]);
                }
            }

            if (frameNames.Count == 0)
            {
                return OperationResult<List<string>>.Failure(
                    "No frame objects are currently selected in the model.");
            }

            return OperationResult<List<string>>.Success(frameNames);
        }

        private static string ObjectTypeIdToName(int typeId)
        {
            switch (typeId)
            {
                case 1:  return "Point";
                case 2:  return "Frame";
                case 5:  return "Shell";
                default: return "Object(" + typeId + ")";
            }
        }
    }
}
