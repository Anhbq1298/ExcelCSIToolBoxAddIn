using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Data.DTOs.CSI;
using ExcelCSIToolBox.Core.Abstractions.CSI;

namespace ExcelCSIToolBox.Core.Application
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

