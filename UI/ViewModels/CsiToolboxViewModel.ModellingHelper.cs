using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;
using ExcelCSIToolBox.Core.Common.Commands;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public partial class CsiToolboxViewModel
    {
        private string _point1Name;
        private string _point2Name;
        private string _referenceFrameName;
        private string _selectedSection;
        private bool _isFixFixSelected;
        private bool _isPinPinSelected;
        private int _numberOfSpaces;
        private bool _isCreateArrayPerpendicularToPathOpen;

        private void InitializeModellingHelperPage()
        {
            AvailableSections = new ObservableCollection<string>();
            _isFixFixSelected = true;
            _numberOfSpaces = 3;

            OpenCreateArrayPerpendicularToPathCommand = new RelayCommand(OpenCreateArrayPerpendicularToPath);
            PickPoint1Command = new RelayCommand(() => ShowPlaceholder("Pick Point 1"));
            PickPoint2Command = new RelayCommand(() => ShowPlaceholder("Pick Point 2"));
            PickReferenceFrameCommand = new RelayCommand(() => ShowPlaceholder("Pick Reference Frame"));
            CreateFramesCommand = new RelayCommand(() => ShowPlaceholder("Create Frames"));
            BackToArrayFrameElementCommand = new RelayCommand(BackToArrayFrameElement);
            CloseCommand = new RelayCommand(BackToArrayFrameElement);
        }

        public string Point1Name
        {
            get { return _point1Name; }
            set
            {
                if (_point1Name == value)
                {
                    return;
                }

                _point1Name = value;
                OnPropertyChanged();
            }
        }

        public string Point2Name
        {
            get { return _point2Name; }
            set
            {
                if (_point2Name == value)
                {
                    return;
                }

                _point2Name = value;
                OnPropertyChanged();
            }
        }

        public string ReferenceFrameName
        {
            get { return _referenceFrameName; }
            set
            {
                if (_referenceFrameName == value)
                {
                    return;
                }

                _referenceFrameName = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> AvailableSections { get; private set; }

        public string SelectedSection
        {
            get { return _selectedSection; }
            set
            {
                if (_selectedSection == value)
                {
                    return;
                }

                _selectedSection = value;
                OnPropertyChanged();
            }
        }

        public bool IsFixFixSelected
        {
            get { return _isFixFixSelected; }
            set
            {
                if (_isFixFixSelected == value)
                {
                    return;
                }

                _isFixFixSelected = value;
                OnPropertyChanged();
                if (value && _isPinPinSelected)
                {
                    _isPinPinSelected = false;
                    OnPropertyChanged(nameof(IsPinPinSelected));
                }
            }
        }

        public bool IsPinPinSelected
        {
            get { return _isPinPinSelected; }
            set
            {
                if (_isPinPinSelected == value)
                {
                    return;
                }

                _isPinPinSelected = value;
                OnPropertyChanged();
                if (value && _isFixFixSelected)
                {
                    _isFixFixSelected = false;
                    OnPropertyChanged(nameof(IsFixFixSelected));
                }
            }
        }

        public int NumberOfSpaces
        {
            get { return _numberOfSpaces; }
            set
            {
                int normalizedValue = value < 1 ? 1 : value;
                if (_numberOfSpaces == normalizedValue)
                {
                    return;
                }

                _numberOfSpaces = normalizedValue;
                OnPropertyChanged();
            }
        }

        public bool IsCreateArrayPerpendicularToPathOpen
        {
            get { return _isCreateArrayPerpendicularToPathOpen; }
            private set
            {
                if (_isCreateArrayPerpendicularToPathOpen == value)
                {
                    return;
                }

                _isCreateArrayPerpendicularToPathOpen = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ArrayFrameElementLandingVisibility));
                OnPropertyChanged(nameof(CreateArrayPerpendicularToPathVisibility));
                OnPropertyChanged(nameof(ActivePageTitle));
                OnPropertyChanged(nameof(ActivePageBreadcrumb));
            }
        }

        public Visibility ArrayFrameElementLandingVisibility
        {
            get { return IsCreateArrayPerpendicularToPathOpen ? Visibility.Collapsed : Visibility.Visible; }
        }

        public Visibility CreateArrayPerpendicularToPathVisibility
        {
            get { return IsCreateArrayPerpendicularToPathOpen ? Visibility.Visible : Visibility.Collapsed; }
        }

        public ICommand OpenCreateArrayPerpendicularToPathCommand { get; private set; }
        public ICommand PickPoint1Command { get; private set; }
        public ICommand PickPoint2Command { get; private set; }
        public ICommand PickReferenceFrameCommand { get; private set; }
        public ICommand CreateFramesCommand { get; private set; }
        public ICommand BackToArrayFrameElementCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }

        private void OpenCreateArrayPerpendicularToPath()
        {
            IsCreateArrayPerpendicularToPathOpen = true;
        }

        private void BackToArrayFrameElement()
        {
            IsCreateArrayPerpendicularToPathOpen = false;
        }
    }
}
