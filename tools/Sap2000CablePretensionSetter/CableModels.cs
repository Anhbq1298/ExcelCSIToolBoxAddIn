using System;
using System.ComponentModel;
using System.Globalization;

namespace Sap2000CablePretensionSetter
{
    internal enum PretensionDefinition
    {
        TensionAtIEnd = 3,
        TensionAtJEnd = 4,
        HorizontalTensionComponent = 5
    }

    internal static class PretensionDefinitionExtensions
    {
        public static string ToDisplayName(this PretensionDefinition definition)
        {
            switch (definition)
            {
                case PretensionDefinition.TensionAtIEnd:
                    return "Tension at I-End";
                case PretensionDefinition.TensionAtJEnd:
                    return "Tension at J-End";
                case PretensionDefinition.HorizontalTensionComponent:
                    return "Horizontal Tension Component";
                default:
                    return definition.ToString();
            }
        }

        public static int ToParameterIndex(this PretensionDefinition definition)
        {
            switch (definition)
            {
                case PretensionDefinition.TensionAtIEnd:
                    return 0;
                case PretensionDefinition.TensionAtJEnd:
                    return 1;
                case PretensionDefinition.HorizontalTensionComponent:
                    return 2;
                default:
                    throw new ArgumentOutOfRangeException(nameof(definition));
            }
        }
    }

    internal sealed class CableDefinitionData
    {
        public int CableType { get; set; }
        public int NumberOfSegments { get; set; }
        public double AddedWeight { get; set; }
        public double ProjectedLoad { get; set; }
        public bool UseDeformedGeometry { get; set; }
        public bool ModelUsingFrames { get; set; }
        public double[] Parameters { get; set; }

        public double GetPretension(PretensionDefinition definition)
        {
            int index = definition.ToParameterIndex();
            if (Parameters == null || index < 0 || index >= Parameters.Length)
            {
                return double.NaN;
            }

            return Parameters[index];
        }

        public string CableTypeName
        {
            get
            {
                switch (CableType)
                {
                    case 1: return "Minimum Tension at I-End";
                    case 2: return "Minimum Tension at J-End";
                    case 3: return "Tension at I-End";
                    case 4: return "Tension at J-End";
                    case 5: return "Horizontal Tension Component";
                    case 6: return "Maximum Vertical Sag";
                    case 7: return "Low-Point Vertical Sag";
                    case 8: return "Undeformed Length";
                    case 9: return "Relative Undeformed Length";
                    default: return "Unknown (" + CableType.ToString(CultureInfo.InvariantCulture) + ")";
                }
            }
        }
    }

    internal sealed class CableInfo
    {
        public CableInfo(string name, string iJoint, string jJoint, CableDefinitionData definition)
        {
            Name = name ?? string.Empty;
            IJoint = iJoint ?? string.Empty;
            JJoint = jJoint ?? string.Empty;
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
        }

        public string Name { get; }
        public string IJoint { get; }
        public string JJoint { get; }
        public CableDefinitionData Definition { get; }
    }

    internal sealed class CableRow : INotifyPropertyChanged
    {
        private bool _isSelected;
        private decimal _currentPretension;
        private decimal _pretensionToSet;
        private string _status;
        private string _currentDefinition;

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public string CableName { get; set; }
        public string IJoint { get; set; }
        public string JJoint { get; set; }

        public string CurrentDefinition
        {
            get { return _currentDefinition; }
            set
            {
                if (_currentDefinition == value) return;
                _currentDefinition = value;
                OnPropertyChanged(nameof(CurrentDefinition));
            }
        }

        public decimal CurrentPretension
        {
            get { return _currentPretension; }
            set
            {
                if (_currentPretension == value) return;
                _currentPretension = value;
                OnPropertyChanged(nameof(CurrentPretension));
            }
        }

        public decimal PretensionToSet
        {
            get { return _pretensionToSet; }
            set
            {
                if (_pretensionToSet == value) return;
                _pretensionToSet = value;
                OnPropertyChanged(nameof(PretensionToSet));
            }
        }

        public string Status
        {
            get { return _status; }
            set
            {
                if (_status == value) return;
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        public CableDefinitionData DefinitionData { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    internal sealed class PretensionLogRow
    {
        public DateTime Time { get; set; }
        public string Cable { get; set; }
        public string Definition { get; set; }
        public decimal Pretension { get; set; }
        public string Result { get; set; }
        public string Message { get; set; }
    }

    internal sealed class OperationResult
    {
        private OperationResult(bool success, string message, int returnCode)
        {
            Success = success;
            Message = message ?? string.Empty;
            ReturnCode = returnCode;
        }

        public bool Success { get; }
        public string Message { get; }
        public int ReturnCode { get; }

        public static OperationResult Ok(string message)
        {
            return new OperationResult(true, message, 0);
        }

        public static OperationResult Fail(string message, int returnCode = -1)
        {
            return new OperationResult(false, message, returnCode);
        }
    }

    internal sealed class OperationResult<T>
    {
        private OperationResult(bool success, T data, string message, int returnCode)
        {
            Success = success;
            Data = data;
            Message = message ?? string.Empty;
            ReturnCode = returnCode;
        }

        public bool Success { get; }
        public T Data { get; }
        public string Message { get; }
        public int ReturnCode { get; }

        public static OperationResult<T> Ok(T data, string message)
        {
            return new OperationResult<T>(true, data, message, 0);
        }

        public static OperationResult<T> Fail(string message, int returnCode = -1)
        {
            return new OperationResult<T>(false, default(T), message, returnCode);
        }
    }
}
