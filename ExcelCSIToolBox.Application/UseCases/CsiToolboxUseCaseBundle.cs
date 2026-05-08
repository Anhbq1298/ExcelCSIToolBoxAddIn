using System;
using ExcelCSIToolBox.Core.Abstractions.CSI;
using ExcelCSIToolBox.Core.Abstractions.Excel;

namespace ExcelCSIToolBox.Application.UseCases
{
    public sealed class CsiToolboxUseCaseBundle
    {
        public CsiToolboxUseCaseBundle(
            ICSISapModelConnectionService csiConnectionService,
            IExcelSelectionService excelSelectionService,
            IExcelOutputService excelOutputService)
        {
            if (csiConnectionService == null) throw new ArgumentNullException(nameof(csiConnectionService));
            if (excelSelectionService == null) throw new ArgumentNullException(nameof(excelSelectionService));
            if (excelOutputService == null) throw new ArgumentNullException(nameof(excelOutputService));

            LoadConnection = new LoadCSISapModelConnectionUseCase(csiConnectionService);
            CloseCurrentInstance = new CloseCurrentInstanceUseCase(csiConnectionService);
            GetSelectedPoints = new GetSelectedCSISapModelPointsUseCase(csiConnectionService, excelOutputService);
            GetSelectedFrames = new GetSelectedCSISapModelFramesUseCase(csiConnectionService, excelOutputService);
            SelectPointsByUniqueName = new SelectPointsFromExcelRangeByUniqueNameUseCase(csiConnectionService, excelSelectionService);
            SelectFramesByUniqueName = new SelectFramesFromExcelRangeByUniqueNameUseCase(csiConnectionService, excelSelectionService);
            AddPointsByCartesian = new AddPointsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            AddFramesByCoordinates = new AddFrameByCoordinatesFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            AddFramesByPointNames = new AddFramesByPointFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateShellAreasFromSelectedFrames = new CreateShellAreasFromSelectedFramesUseCase(csiConnectionService);

            CreateSteelISections = new CreateSteelISectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateSteelChannelSections = new CreateSteelChannelSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateSteelAngleSections = new CreateSteelAngleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateSteelPipeSections = new CreateSteelPipeSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateSteelTubeSections = new CreateSteelTubeSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);

            CreateConcreteRectangleSections = new CreateConcreteRectangleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);
            CreateConcreteCircleSections = new CreateConcreteCircleSectionsFromExcelRangeUseCase(csiConnectionService, excelSelectionService);

            GetLoadCombinations = new GetLoadCombinationsUseCase(csiConnectionService);
            DeleteLoadCombinations = new DeleteLoadCombinationsUseCase(csiConnectionService);
            GetLoadCombinationDetails = new GetLoadCombinationDetailsUseCase(csiConnectionService);

            GetLoadPatterns = new GetLoadPatternsUseCase(csiConnectionService);
            DeleteLoadPatterns = new DeleteLoadPatternsUseCase(csiConnectionService);

            GetFrameSections = new GetFrameSectionsUseCase(csiConnectionService);
            GetFrameSectionDetail = new GetFrameSectionDetailUseCase(csiConnectionService);
            UpdateFrameSection = new UpdateFrameSectionUseCase(csiConnectionService);
            RenameFrameSection = new RenameFrameSectionUseCase(csiConnectionService);
        }

        public LoadCSISapModelConnectionUseCase LoadConnection { get; private set; }
        public CloseCurrentInstanceUseCase CloseCurrentInstance { get; private set; }
        public GetSelectedCSISapModelPointsUseCase GetSelectedPoints { get; private set; }
        public GetSelectedCSISapModelFramesUseCase GetSelectedFrames { get; private set; }
        public SelectPointsFromExcelRangeByUniqueNameUseCase SelectPointsByUniqueName { get; private set; }
        public SelectFramesFromExcelRangeByUniqueNameUseCase SelectFramesByUniqueName { get; private set; }
        public AddPointsFromExcelRangeUseCase AddPointsByCartesian { get; private set; }
        public AddFrameByCoordinatesFromExcelRangeUseCase AddFramesByCoordinates { get; private set; }
        public AddFramesByPointFromExcelRangeUseCase AddFramesByPointNames { get; private set; }
        public CreateShellAreasFromSelectedFramesUseCase CreateShellAreasFromSelectedFrames { get; private set; }
        public CreateSteelISectionsFromExcelRangeUseCase CreateSteelISections { get; private set; }
        public CreateSteelChannelSectionsFromExcelRangeUseCase CreateSteelChannelSections { get; private set; }
        public CreateSteelAngleSectionsFromExcelRangeUseCase CreateSteelAngleSections { get; private set; }
        public CreateSteelPipeSectionsFromExcelRangeUseCase CreateSteelPipeSections { get; private set; }
        public CreateSteelTubeSectionsFromExcelRangeUseCase CreateSteelTubeSections { get; private set; }
        public CreateConcreteRectangleSectionsFromExcelRangeUseCase CreateConcreteRectangleSections { get; private set; }
        public CreateConcreteCircleSectionsFromExcelRangeUseCase CreateConcreteCircleSections { get; private set; }
        public GetLoadCombinationsUseCase GetLoadCombinations { get; private set; }
        public DeleteLoadCombinationsUseCase DeleteLoadCombinations { get; private set; }
        public GetLoadCombinationDetailsUseCase GetLoadCombinationDetails { get; private set; }
        public GetLoadPatternsUseCase GetLoadPatterns { get; private set; }
        public DeleteLoadPatternsUseCase DeleteLoadPatterns { get; private set; }
        public GetFrameSectionsUseCase GetFrameSections { get; private set; }
        public GetFrameSectionDetailUseCase GetFrameSectionDetail { get; private set; }
        public UpdateFrameSectionUseCase UpdateFrameSection { get; private set; }
        public RenameFrameSectionUseCase RenameFrameSection { get; private set; }
    }
}
