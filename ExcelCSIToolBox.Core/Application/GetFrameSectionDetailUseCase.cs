using System;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
{
    public class GetFrameSectionDetailUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetFrameSectionDetailUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public OperationResult<CSISapModelFrameSectionDetailDTO> Execute(string sectionName)
        {
            return _connectionService.GetFrameSectionDetail(sectionName);
        }
    }
}

