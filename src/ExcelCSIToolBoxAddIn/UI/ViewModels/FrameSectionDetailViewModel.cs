using System.Collections.ObjectModel;
using System.Globalization;
using ExcelCSIToolBox.Data.DTOs.CSI;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public class FrameSectionDetailViewModel : ViewModelBase
    {
        private string _sectionName;
        private string _materialName;

        public FrameSectionDetailViewModel(CSISapModelFrameSectionDetailDTO detail)
        {
            OriginalName = detail.Name;
            SectionName = detail.Name;
            MaterialName = detail.MaterialName;
            ShapeType = detail.ShapeType;
            Color = detail.Color;
            Notes = detail.Notes;
            Dimensions = new ObservableCollection<FrameSectionDimensionEditItem>();

            foreach (var pair in detail.Dimensions)
            {
                Dimensions.Add(new FrameSectionDimensionEditItem
                {
                    Key = pair.Key,
                    ValueText = pair.Value.ToString("0.###", CultureInfo.InvariantCulture)
                });
            }
        }

        public string OriginalName { get; }

        public string SectionName
        {
            get { return _sectionName; }
            set
            {
                _sectionName = value;
                OnPropertyChanged();
            }
        }

        public string MaterialName
        {
            get { return _materialName; }
            set
            {
                _materialName = value;
                OnPropertyChanged();
            }
        }

        public FrameSectionShapeType ShapeType { get; }
        public int Color { get; }
        public string Notes { get; }
        public ObservableCollection<FrameSectionDimensionEditItem> Dimensions { get; }

        public bool IsRename => !string.Equals(OriginalName, SectionName, System.StringComparison.Ordinal);

        public CSISapModelFrameSectionUpdateDTO ToUpdateDto()
        {
            var dto = new CSISapModelFrameSectionUpdateDTO
            {
                OriginalName = OriginalName,
                SectionName = SectionName?.Trim(),
                MaterialName = MaterialName?.Trim(),
                ShapeType = ShapeType,
                Color = Color,
                Notes = Notes
            };

            foreach (var item in Dimensions)
            {
                if (double.TryParse(item.ValueText, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.TryParse(item.ValueText, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
                {
                    dto.Dimensions[item.Key] = value;
                }
            }

            return dto;
        }

        public CSISapModelFrameSectionRenameDTO ToRenameDto()
        {
            var update = ToUpdateDto();
            return new CSISapModelFrameSectionRenameDTO
            {
                OriginalName = update.OriginalName,
                SectionName = update.SectionName,
                MaterialName = update.MaterialName,
                ShapeType = update.ShapeType,
                Dimensions = update.Dimensions,
                Color = update.Color,
                Notes = update.Notes
            };
        }
    }
}

