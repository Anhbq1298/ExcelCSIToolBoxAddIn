using System;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data.DTOs;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
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
