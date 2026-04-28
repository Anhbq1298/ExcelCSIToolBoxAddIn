using System;
using System.Collections.Generic;
using ExcelCSIToolBoxAddIn.Common.Results;
using ExcelCSIToolBoxAddIn.Data;
using ExcelCSIToolBoxAddIn.Infrastructure.CSISapModel;

namespace ExcelCSIToolBoxAddIn.Core.Application
{
    public class GetFrameSectionsUseCase
    {
        private readonly ICSISapModelConnectionService _connectionService;

        public GetFrameSectionsUseCase(ICSISapModelConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public OperationResult<IReadOnlyList<CSISapModelFrameSectionDTO>> Execute()
        {
            return _connectionService.GetFrameSections();
        }
    }
}
