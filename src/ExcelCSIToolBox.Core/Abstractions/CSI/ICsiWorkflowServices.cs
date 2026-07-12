using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.CSISapModel.Random;
using ExcelCSIToolBox.Data.CSISapModel.Truss;
using ExcelCSIToolBox.Data.CSISapModel.Workflow;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ICsiRandomObjectGenerationService
    {
        OperationResult<RandomCsiObjectResultDto> Generate(
            ICSISapModelConnectionService service,
            RandomCsiObjectRequestDto request);
    }

    public interface ICsiTrussGenerationService
    {
        OperationResult<HoweTrussResultDto> Generate(
            ICSISapModelConnectionService service,
            HoweTrussRequestDto request);
    }

    public interface ICsiWorkflowExecutionService
    {
        OperationResult<CsiWorkflowResultDto> Execute(
            ICSISapModelConnectionService service,
            CsiWorkflowRequestDto request);
    }
}
