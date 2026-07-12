using System.Collections.Generic;
using ExcelCSIToolBox.Core.Common.Results;
using ExcelCSIToolBox.Core.Models.CSI;

namespace ExcelCSIToolBox.Core.Abstractions.CSI
{
    public interface ISelectedObjectIdentityResolver
    {
        OperationResult<IReadOnlyList<CsiObjectIdentity>> ResolveSelectedObjects();
    }
}
