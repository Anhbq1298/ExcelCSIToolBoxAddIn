using System;
using System.Collections.Generic;
using System.IO;
using ExcelCSIToolBox.RefBuilder.Chm;
using ExcelCSIToolBox.RefBuilder.Generation;
using ExcelCSIToolBox.RefBuilder.Indexing;
using ExcelCSIToolBox.RefBuilder.Parsing;

namespace ExcelCSIToolBox.RefBuilder
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            string root = args.Length > 0 ? args[0] : FindRepositoryRoot();
            string refRoot = Path.Combine(root, "_ref");
            string infrastructureRoot = Path.Combine(root, "ExcelCSIToolBox.Infrastructure");

            IChmExtractor extractor = new ChmExtractor();
            IApiDocParser parser = new ReflectionApiDocParser();
            IApiIndexBuilder builder = new ApiIndexBuilder(parser);
            ApiIndexWriter writer = new ApiIndexWriter();
            IServiceScaffoldGenerator generator = new ServiceScaffoldGenerator();

            extractor.Extract(refRoot);

            IReadOnlyList<ApiMethodDefinition> etabs = builder.Build("ETABS", Path.Combine(root, "ETABSv1.dll"));
            IReadOnlyList<ApiMethodDefinition> sap2000 = builder.Build("SAP2000", Path.Combine(root, "SAP2000v1.dll"));

            writer.Write(Path.Combine(refRoot, "ETABS", "index", "etabs_api_index.json"), etabs);
            writer.Write(Path.Combine(refRoot, "SAP2000", "index", "sap2000_api_index.json"), sap2000);

            List<ApiMethodDefinition> all = new List<ApiMethodDefinition>();
            all.AddRange(etabs);
            all.AddRange(sap2000);
            generator.Generate(infrastructureRoot, all);

            Console.WriteLine("CSI reference index generation complete.");
            Console.WriteLine("ETABS methods: " + etabs.Count);
            Console.WriteLine("SAP2000 methods: " + sap2000.Count);
            return 0;
        }

        private static string FindRepositoryRoot()
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            while (!string.IsNullOrWhiteSpace(directory))
            {
                if (File.Exists(Path.Combine(directory, "ExcelCSIToolBoxAddIn.sln")))
                {
                    return directory;
                }

                DirectoryInfo parent = Directory.GetParent(directory);
                directory = parent == null ? null : parent.FullName;
            }

            return Directory.GetCurrentDirectory();
        }
    }
}
