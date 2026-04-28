using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel.ReadOnly
{
    /// <summary>
    /// Read-only frame service that reads section assignments from selected frames.
    /// Calls FrameObj.GetSection on the active CSI model without modifying anything.
    /// Raw SapModel objects are kept entirely inside this class.
    /// </summary>
    public class CsiReadOnlyFrameService : ICsiReadOnlyFrameService
    {
        private readonly ICSISapModelConnectionAdapter<ETABSv1.cSapModel>   _etabsAdapter;
        private readonly ICSISapModelConnectionAdapter<SAP2000v1.cSapModel> _sap2000Adapter;

        public CsiReadOnlyFrameService()
        {
            _etabsAdapter   = CSISapModelConnectionAdapterFactory.CreateEtabs();
            _sap2000Adapter = CSISapModelConnectionAdapterFactory.CreateSap2000();
        }

        /// <summary>
        /// Return the section assignments for all currently selected frame objects.
        /// Tries ETABS first, then SAP2000. Does not modify the model.
        /// </summary>
        public OperationResult<List<FrameSectionAssignmentDto>> GetSelectedFrameSections()
        {
            // Try ETABS.
            OperationResult<ETABSv1.cSapModel> etabsModel = _etabsAdapter.EnsureSapModel();
            if (etabsModel.IsSuccess)
            {
                return GetSectionsFromEtabs(etabsModel.Data);
            }

            // Try SAP2000.
            OperationResult<SAP2000v1.cSapModel> sap2000Model = _sap2000Adapter.EnsureSapModel();
            if (sap2000Model.IsSuccess)
            {
                return GetSectionsFromSap2000(sap2000Model.Data);
            }

            return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                "No running ETABS or SAP2000 instance found. Please open a model first.");
        }

        // ─── ETABS implementation ────────────────────────────────────────────────

        private static OperationResult<List<FrameSectionAssignmentDto>> GetSectionsFromEtabs(
            ETABSv1.cSapModel sapModel)
        {
            try
            {
                // First get all selected frames.
                int      numberItems = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int selRet = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (selRet != 0)
                {
                    return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                        "ETABS SelectObj.GetSelected returned an error.");
                }

                List<string> selectedFrames = ExtractFrameNames(numberItems, objectTypes, objectNames);
                if (selectedFrames.Count == 0)
                {
                    return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                        "No frame objects are currently selected in the ETABS model.");
                }

                // Now get the section for each selected frame.
                var assignments = new List<FrameSectionAssignmentDto>();
                foreach (string frameName in selectedFrames)
                {
                    string propName = string.Empty;
                    string autoName = string.Empty;
                    int getSecRet = sapModel.FrameObj.GetSection(frameName, ref propName, ref autoName);
                    assignments.Add(new FrameSectionAssignmentDto
                    {
                        FrameName   = frameName,
                        SectionName = getSecRet == 0 ? propName : "(unavailable)"
                    });
                }

                return OperationResult<List<FrameSectionAssignmentDto>>.Success(assignments);
            }
            catch (Exception ex)
            {
                return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                    "Failed to read frame sections from ETABS: " + ex.Message);
            }
        }

        // ─── SAP2000 implementation ──────────────────────────────────────────────

        private static OperationResult<List<FrameSectionAssignmentDto>> GetSectionsFromSap2000(
            SAP2000v1.cSapModel sapModel)
        {
            try
            {
                int      numberItems = 0;
                int[]    objectTypes = null;
                string[] objectNames = null;

                int selRet = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
                if (selRet != 0)
                {
                    return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                        "SAP2000 SelectObj.GetSelected returned an error.");
                }

                List<string> selectedFrames = ExtractFrameNames(numberItems, objectTypes, objectNames);
                if (selectedFrames.Count == 0)
                {
                    return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                        "No frame objects are currently selected in the SAP2000 model.");
                }

                var assignments = new List<FrameSectionAssignmentDto>();
                foreach (string frameName in selectedFrames)
                {
                    string propName = string.Empty;
                    string autoName = string.Empty;
                    int getSecRet = sapModel.FrameObj.GetSection(frameName, ref propName, ref autoName);
                    assignments.Add(new FrameSectionAssignmentDto
                    {
                        FrameName   = frameName,
                        SectionName = getSecRet == 0 ? propName : "(unavailable)"
                    });
                }

                return OperationResult<List<FrameSectionAssignmentDto>>.Success(assignments);
            }
            catch (Exception ex)
            {
                return OperationResult<List<FrameSectionAssignmentDto>>.Failure(
                    "Failed to read frame sections from SAP2000: " + ex.Message);
            }
        }

        // ─── Shared helpers ──────────────────────────────────────────────────────

        private static List<string> ExtractFrameNames(
            int numberItems, int[] objectTypes, string[] objectNames)
        {
            var names = new List<string>();
            for (int i = 0; i < numberItems; i++)
            {
                if (objectTypes == null || objectNames == null ||
                    i >= objectTypes.Length || i >= objectNames.Length)
                {
                    continue;
                }

                // Object type 2 = Frame.
                if (objectTypes[i] == 2 && !string.IsNullOrWhiteSpace(objectNames[i]))
                {
                    names.Add(objectNames[i]);
                }
            }

            return names;
        }
    }
}
