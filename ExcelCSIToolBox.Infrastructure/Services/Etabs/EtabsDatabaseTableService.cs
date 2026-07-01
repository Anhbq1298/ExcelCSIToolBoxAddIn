using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ETABSv1;
using ExcelCSIToolBox.Application.Interfaces.Etabs;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs
{
    public class EtabsDatabaseTableService : IEtabsDatabaseTableService
    {
        private readonly IEtabsConnectionService _connectionService;

        public EtabsDatabaseTableService(IEtabsConnectionService connectionService)
        {
            _connectionService = connectionService ?? throw new ArgumentNullException(nameof(connectionService));
        }

        public Task<EtabsTableResult> GetTableAsync(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("ETABS table name is required.", nameof(tableName));
            }

            return Task.Run(() =>
            {
                cSapModel sapModel = _connectionService.SapModel as cSapModel;
                if (sapModel == null)
                {
                    throw new InvalidOperationException("The attached ETABS model is invalid. Please reattach and try again.");
                }

                int tableVersion = 0;
                string[] fieldKeyList = null;
                string[] fieldsKeysIncluded = null;
                int numberRecords = 0;
                string[] tableData = null;

                int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    tableName,
                    ref fieldKeyList,
                    string.Empty,
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData);

                if (ret != 0)
                {
                    throw new InvalidOperationException("Failed to extract ETABS table '" + tableName + "' (return code " + ret + ").");
                }

                EtabsTableResult result = new EtabsTableResult { TableName = tableName };
                string[] returnedFields = fieldsKeysIncluded != null && fieldsKeysIncluded.Length > 0
                    ? fieldsKeysIncluded
                    : fieldKeyList;

                if (returnedFields != null)
                {
                    result.Headers.AddRange(returnedFields);
                }

                int columnCount = result.Headers.Count;
                if (tableData != null && columnCount > 0)
                {
                    for (int i = 0; i < numberRecords; i++)
                    {
                        List<string> row = new List<string>();
                        for (int j = 0; j < columnCount; j++)
                        {
                            int index = i * columnCount + j;
                            row.Add(index < tableData.Length ? tableData[index] : string.Empty);
                        }

                        result.Rows.Add(row);
                    }
                }

                return result;
            });
        }
    }
}
