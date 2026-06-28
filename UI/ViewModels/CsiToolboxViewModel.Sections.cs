using System;
using System.Windows;
using System.Globalization;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private void RefreshSectionDimensionAnnotations()
        {
            SectionDimensionAnnotations.Clear();

            if (SelectedFrameSectionDetail == null)
            {
                return;
            }

            foreach (var annotation in CreateDimensionAnnotations(SelectedFrameSectionDetail, GetLengthUnitText()))
            {
                SectionDimensionAnnotations.Add(annotation);
            }
        }

        private static System.Collections.Generic.IEnumerable<SectionDimensionAnnotation> CreateDimensionAnnotations(CSISapModelFrameSectionDetailDTO detail, string unit)
        {
            switch (detail.ShapeType)
            {
                case FrameSectionShapeType.Rectangular:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Depth ( t3 )", "Total depth ( t3 )"),
                        Spec("b", "Width ( t2 )", "Flange width ( t2 )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Tube:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Flange width ( t2 )", "Width ( t2 )"),
                        Spec("t2", "Flange thickness ( tf )"),
                        Spec("t3", "Web thickness ( tw )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.I:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Top flange width ( t2 )", "Flange width ( t2 )"),
                        Spec("tw", "Web thickness ( tw )"),
                        Spec("tf", "Top flange thickness ( tf )", "Flange thickness ( tf )"),
                        Spec("t2b", "Bottom flange width ( t2b )"),
                        Spec("tfb", "Bottom flange thickness ( tfb )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Channel:
                case FrameSectionShapeType.Angle:
                case FrameSectionShapeType.DoubleAngle:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Flange width ( t2 )", "Width ( t2 )"),
                        Spec("tw", "Web thickness ( tw )"),
                        Spec("tf", "Flange thickness ( tf )"),
                        Spec("dis", "Spacing ( dis )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Pipe:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("d", "Outside diameter ( t3 )", "Diameter ( t3 )"),
                        Spec("t", "Wall thickness ( tw )", "Wall thickness ( t )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.Circular:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("d", "Diameter ( t3 )", "Outside diameter ( t3 )"),
                        Spec("r", "Radius ( r )")))
                    {
                        yield return item;
                    }
                    break;

                case FrameSectionShapeType.General:
                    foreach (var item in DimensionItems(detail, unit,
                        Spec("h", "Total depth ( t3 )", "Depth ( t3 )"),
                        Spec("b", "Width ( t2 )")))
                    {
                        yield return item;
                    }
                    break;
            }
        }

        private static System.Collections.Generic.IEnumerable<SectionDimensionAnnotation> DimensionItems(
            CSISapModelFrameSectionDetailDTO detail,
            string unit,
            params DimensionSpec[] specs)
        {
            foreach (var spec in specs)
            {
                if (TryGetDimensionValue(detail, out double value, spec.DimensionNames))
                {
                    yield return CreateDimensionItem(spec.Key, value, unit, detail.ShapeType.ToString());
                }
            }
        }

        private static SectionDimensionAnnotation CreateDimensionItem(string key, double value, string unit, string sectionType)
        {
            string valueText = value.ToString("0.###", CultureInfo.InvariantCulture);
            string displayText = string.IsNullOrWhiteSpace(unit)
                ? $"{key} = {valueText}"
                : $"{key} = {valueText} {unit}";

            return new SectionDimensionAnnotation
            {
                Key = key,
                DisplayLabel = key,
                Value = value,
                Unit = unit,
                DisplayText = displayText,
                DescriptionText = $"{key} = {GetDimensionDescription(key, sectionType)}",
                SectionType = sectionType
            };
        }

        private static string GetDimensionDescription(string key, string sectionType)
        {
            switch (key)
            {
                case "h": return "height";
                case "b": return "width";
                case "d": return "diameter";
                case "r": return "radius";
                case "t": return "thickness";
                case "tw": return "web thickness";
                case "tf": return "flange thickness";
                case "t2": return sectionType == FrameSectionShapeType.Tube.ToString() ? "top/bottom wall thickness" : "local 2 dimension";
                case "t3": return sectionType == FrameSectionShapeType.Tube.ToString() ? "side wall thickness" : "local 3 dimension";
                case "t2b": return "bottom flange width";
                case "tfb": return "bottom flange thickness";
                case "dis": return "spacing";
                default: return "dimension";
            }
        }

        private static DimensionSpec Spec(string key, params string[] dimensionNames)
        {
            return new DimensionSpec { Key = key, DimensionNames = dimensionNames };
        }

        private static bool TryGetDimensionValue(CSISapModelFrameSectionDetailDTO detail, out double value, params string[] keys)
        {
            value = 0;
            if (detail?.Dimensions == null)
            {
                return false;
            }

            foreach (string key in keys)
            {
                if (detail.Dimensions.TryGetValue(key, out value))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetLengthUnitText()
        {
            string unitText = CurrentModelUnitText ?? string.Empty;
            string lower = unitText.ToLowerInvariant();

            if (lower.Contains("mm")) return "mm";
            if (lower.Contains("cm")) return "cm";
            if (lower.Contains("-m-") || lower.EndsWith("-m")) return "m";
            if (lower.Contains("in")) return "in";
            if (lower.Contains("ft")) return "ft";

            return string.Empty;
        }

        private class DimensionSpec
         {
             public string Key { get; set; }
             public string[] DimensionNames { get; set; }
         }

        private void LoadSelectedSectionDetail(CSISapModelFrameSectionDTO section)
        {
            if (section == null || _useCases.GetFrameSectionDetail == null)
            {
                SelectedFrameSectionDetail = null;
                return;
            }
            var result = _useCases.GetFrameSectionDetail.Execute(section.Name);
            SelectedFrameSectionDetail = result.IsSuccess ? result.Data : null;
        }

        private void GetFrameSections()
        {
            var result = _useCases.GetFrameSections.Execute();
            if (result.IsSuccess)
            {
                FrameSections.Clear();
                if (result.Data != null)
                {
                    foreach (var section in result.Data)
                    {
                        FrameSections.Add(section);
                    }
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void EditFrameSection(CSISapModelFrameSectionDTO section)
        {
            if (section == null) return;

            var result = _useCases.GetFrameSectionDetail.Execute(section.Name);
            if (result.IsSuccess)
            {
                var window = new ExcelCSIToolBoxAddIn.UI.Views.FrameSectionDetailWindow(result.Data);
                bool? dialogResult = window.ShowDialog();
                if (dialogResult != true)
                {
                    return;
                }

                OperationResult saveResult;
                string selectedName;
                if (window.ViewModel.IsRename)
                {
                    var confirm = MessageBox.Show(
                        "Renaming a section will create a new section, reassign frames using the old section, and then delete the old section when possible. Continue?",
                        ProductTitle,
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Warning);

                    if (confirm != MessageBoxResult.OK)
                    {
                        return;
                    }

                    var renameInput = window.ViewModel.ToRenameDto();
                    selectedName = renameInput.SectionName;
                    saveResult = _useCases.RenameFrameSection.Execute(renameInput);
                }
                else
                {
                    var updateInput = window.ViewModel.ToUpdateDto();
                    selectedName = updateInput.SectionName;
                    saveResult = _useCases.UpdateFrameSection.Execute(updateInput);
                }

                ShowOperationResult(saveResult);
                if (saveResult.IsSuccess)
                {
                    GetFrameSections();
                    SelectFrameSectionByName(selectedName);
                }
            }
            else
            {
                ShowOperationResult(OperationResult.Failure(result.Message));
            }
        }

        private void SelectFrameSectionByName(string sectionName)
        {
            foreach (var section in FrameSections)
            {
                if (string.Equals(section.Name, sectionName, System.StringComparison.Ordinal))
                {
                    SelectedFrameSection = section;
                    return;
                }
            }
        }
    }
}
