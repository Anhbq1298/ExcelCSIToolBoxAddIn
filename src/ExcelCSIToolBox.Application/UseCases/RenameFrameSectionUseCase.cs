using System;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Application.UseCases
{
    public class RenameFrameSectionUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public RenameFrameSectionUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public OperationResult Execute(CSISapModelFrameSectionRenameDTO input)
        {
            return _connectionService.RenameFrameSection(input);
        }
    }
}

