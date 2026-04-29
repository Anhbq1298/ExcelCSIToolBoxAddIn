using System;
using ExcelCSIToolBox.Core.Common.Results;

namespace ExcelCSIToolBox.Infrastructure.CSISapModel
{
    internal delegate int CSISapModelClearSelection<TSapModel>(TSapModel sapModel);

    internal delegate int CSISapModelSetSelectedByName<TSapModel>(
        TSapModel sapModel,
        string objectName);

    internal delegate int CSISapModelGetSelectedObjects<TSapModel>(
        TSapModel sapModel,
        ref int numberItems,
        ref int[] objectTypes,
        ref string[] objectNames);

    internal delegate int CSISapModelReadCount<TSapModel>(
        TSapModel sapModel,
        ref int count);

    internal delegate int CSISapModelGetNameList<TSapModel>(
        TSapModel sapModel,
        ref int numberNames,
        ref string[] names);

    internal delegate int CSISapModelGetPointCoordinates<TSapModel>(
        TSapModel sapModel,
        string pointName,
        ref double x,
        ref double y,
        ref double z);
}
