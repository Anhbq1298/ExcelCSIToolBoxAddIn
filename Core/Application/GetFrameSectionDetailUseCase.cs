using System;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
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
