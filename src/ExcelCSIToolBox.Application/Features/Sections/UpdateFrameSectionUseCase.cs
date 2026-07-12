using System;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Contracts.CSI;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.Features.Sections
{
    public class UpdateFrameSectionUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public UpdateFrameSectionUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public OperationResult Execute(CSISapModelFrameSectionUpdateDTO input)
        {
            return _connectionService.UpdateFrameSection(input);
        }
    }
}

