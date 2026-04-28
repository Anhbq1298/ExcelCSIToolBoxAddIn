using System;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
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
