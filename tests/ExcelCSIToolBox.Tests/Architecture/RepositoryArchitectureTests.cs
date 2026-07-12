using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace ExcelCSIToolBox.Tests.Architecture
{
    public class RepositoryArchitectureTests
    {
        [Fact]
        public void Projects_DoNotUseLinkedCompileItems()
        {
            List<string> failures = new List<string>();

            foreach (string projectFile in EnumerateProjectFiles())
            {
                XDocument project = XDocument.Load(projectFile);
                foreach (XElement compile in Elements(project, "Compile"))
                {
                    XAttribute include = compile.Attribute("Include");
                    if (include == null)
                    {
                        continue;
                    }

                    string value = include.Value.Replace('/', '\\');
                    if (value.StartsWith("..\\", StringComparison.Ordinal) ||
                        value.Contains("\\..\\"))
                    {
                        failures.Add(ToRelativePath(projectFile) + " links compile item outside project: " + include.Value);
                    }
                }
            }

            AssertNoFailures(failures);
        }

        [Fact]
        public void ProjectReferences_FollowLayeringRules()
        {
            var expectedReferences = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["src/ExcelCSIToolBox.Core/ExcelCSIToolBox.Core.csproj"] = new string[0],
                ["src/ExcelCSIToolBox.Application/ExcelCSIToolBox.Application.csproj"] = new[] { "ExcelCSIToolBox.Core" },
                ["src/ExcelCSIToolBox.Infrastructure/ExcelCSIToolBox.Infrastructure.csproj"] = new[] { "ExcelCSIToolBox.Application", "ExcelCSIToolBox.Core" },
                ["src/ExcelCSIToolBox.AI/ExcelCSIToolBox.AI.csproj"] = new[] { "ExcelCSIToolBox.Application", "ExcelCSIToolBox.Core" },
                ["src/ExcelCSIToolBoxAddIn/ExcelCSIToolBoxAddIn.csproj"] = new[] { "ExcelCSIToolBox.AI", "ExcelCSIToolBox.Application", "ExcelCSIToolBox.Core", "ExcelCSIToolBox.Infrastructure" },
                ["tests/ExcelCSIToolBox.Tests/ExcelCSIToolBox.Tests.csproj"] = new[] { "ExcelCSIToolBox.Application", "ExcelCSIToolBox.Core", "ExcelCSIToolBox.Infrastructure" }
            };

            List<string> failures = new List<string>();
            foreach (KeyValuePair<string, string[]> expected in expectedReferences)
            {
                string projectFile = Path.Combine(RepositoryRoot, NormalizePath(expected.Key));
                string[] actual = GetProjectReferences(projectFile).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
                string[] allowed = expected.Value.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
                if (!actual.SequenceEqual(allowed, StringComparer.OrdinalIgnoreCase))
                {
                    failures.Add(expected.Key + " references [" + string.Join(", ", actual) + "] but expected [" + string.Join(", ", allowed) + "]");
                }
            }

            AssertNoFailures(failures);
        }

        [Fact]
        public void CoreApplicationAndAi_DoNotReferenceUiInfrastructureOrInterop()
        {
            var projects = new[]
            {
                "src/ExcelCSIToolBox.Core",
                "src/ExcelCSIToolBox.Application",
                "src/ExcelCSIToolBox.AI"
            };

            var forbiddenTokens = new[]
            {
                "using ETABSv1;",
                "using SAP2000v1;",
                "using Microsoft.Office.Interop",
                "using System.Windows;",
                "using System.Windows.Controls;",
                "using System.Windows.Forms;",
                "PresentationFramework",
                "PresentationCore",
                "WindowsBase"
            };

            List<string> failures = new List<string>();
            foreach (string project in projects)
            {
                string projectPath = Path.Combine(RepositoryRoot, NormalizePath(project));
                foreach (string sourceFile in Directory.GetFiles(projectPath, "*.cs", SearchOption.AllDirectories))
                {
                    string text = File.ReadAllText(sourceFile);
                    foreach (string token in forbiddenTokens)
                    {
                        if (text.Contains(token))
                        {
                            failures.Add(ToRelativePath(sourceFile) + " contains forbidden dependency token: " + token);
                        }
                    }
                }

                foreach (string projectFile in Directory.GetFiles(projectPath, "*.csproj", SearchOption.TopDirectoryOnly))
                {
                    string text = File.ReadAllText(projectFile);
                    foreach (string token in forbiddenTokens)
                    {
                        if (text.Contains(token))
                        {
                            failures.Add(ToRelativePath(projectFile) + " contains forbidden dependency token: " + token);
                        }
                    }
                }
            }

            AssertNoFailures(failures);
        }

        [Fact]
        public void DirectComInteropUsings_AreConfinedToAdapterOrHostProjects()
        {
            List<string> failures = new List<string>();
            foreach (string sourceFile in Directory.GetFiles(Path.Combine(RepositoryRoot, "src"), "*.cs", SearchOption.AllDirectories))
            {
                string relativePath = ToRelativePath(sourceFile);
                string[] lines = File.ReadAllLines(sourceFile);
                foreach (string line in lines)
                {
                    string trimmed = line.Trim();
                    if (trimmed == "using ETABSv1;" && !IsUnder(relativePath, "src\\ExcelCSIToolBox.Infrastructure\\CSI\\Etabs\\"))
                    {
                        failures.Add(relativePath + " uses ETABSv1 outside the ETABS adapter boundary.");
                    }

                    if (trimmed == "using SAP2000v1;" && !IsUnder(relativePath, "src\\ExcelCSIToolBox.Infrastructure\\CSI\\Sap2000\\"))
                    {
                        failures.Add(relativePath + " uses SAP2000v1 outside the SAP2000 adapter boundary.");
                    }

                    if (trimmed == "using Microsoft.Office.Interop.Excel;" &&
                        !IsUnder(relativePath, "src\\ExcelCSIToolBox.Infrastructure\\Excel\\") &&
                        !IsUnder(relativePath, "src\\ExcelCSIToolBoxAddIn\\"))
                    {
                        failures.Add(relativePath + " uses Excel interop outside Infrastructure/Excel or the AddIn host.");
                    }
                }
            }

            AssertNoFailures(failures);
        }

        [Fact]
        public void ObsoleteFolders_AreAbsent()
        {
            var obsoletePaths = new[]
            {
                "src/ExcelCSIToolBox.Data",
                "src/ExcelCSIToolBox.Infrastructure/CSISapModel",
                "src/ExcelCSIToolBox.Infrastructure/Etabs",
                "src/ExcelCSIToolBox.Infrastructure/Sap2000",
                "src/ExcelCSIToolBox.Infrastructure/Services/Etabs",
                "src/ExcelCSIToolBox.Infrastructure/Services/Excel",
                "src/ExcelCSIToolBoxAddIn/UI/Views",
                "src/ExcelCSIToolBoxAddIn/UI/ViewModels",
                "src/ExcelCSIToolBoxAddIn/UI/Commands",
                "src/ExcelCSIToolBoxAddIn/UI/Converters",
                "src/ExcelCSIToolBoxAddIn/AddIn/AiTaskPaneManager.cs",
                "src/ExcelCSIToolBoxAddIn/AddIn/AiChatSessionService.cs",
                "src/ExcelCSIToolBoxAddIn/AddIn/BatchProgressReporter.cs",
                "src/ExcelCSIToolBoxAddIn/AddIn/WindowManager.cs"
            };

            List<string> failures = obsoletePaths
                .Where(path => Directory.Exists(Path.Combine(RepositoryRoot, NormalizePath(path))) ||
                               File.Exists(Path.Combine(RepositoryRoot, NormalizePath(path))))
                .ToList();

            AssertNoFailures(failures);
        }

        private static IEnumerable<string> EnumerateProjectFiles()
        {
            return Directory.GetFiles(RepositoryRoot, "*.csproj", SearchOption.AllDirectories)
                .Where(path => !IsUnder(ToRelativePath(path), "bin\\") &&
                               !IsUnder(ToRelativePath(path), "obj\\"));
        }

        private static IEnumerable<XElement> Elements(XDocument document, string localName)
        {
            return document.Descendants().Where(x => x.Name.LocalName == localName);
        }

        private static IEnumerable<string> GetProjectReferences(string projectFile)
        {
            XDocument project = XDocument.Load(projectFile);
            return Elements(project, "ProjectReference")
                .Select(reference => reference.Attribute("Include"))
                .Where(attribute => attribute != null)
                .Select(attribute => Path.GetFileNameWithoutExtension(attribute.Value.Replace('/', Path.DirectorySeparatorChar)))
                .Where(name => !string.IsNullOrWhiteSpace(name));
        }

        private static bool IsUnder(string relativePath, string expectedPrefix)
        {
            return relativePath.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizePath(string relativePath)
        {
            return relativePath.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string ToRelativePath(string fullPath)
        {
            Uri root = new Uri(EnsureTrailingSeparator(RepositoryRoot));
            Uri path = new Uri(Path.GetFullPath(fullPath));
            return Uri.UnescapeDataString(root.MakeRelativeUri(path).ToString()).Replace('/', '\\');
        }

        private static string EnsureTrailingSeparator(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? path
                : path + Path.DirectorySeparatorChar;
        }

        private static void AssertNoFailures(IReadOnlyCollection<string> failures)
        {
            Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
        }

        private static string RepositoryRoot
        {
            get
            {
                DirectoryInfo current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                while (current != null)
                {
                    if (File.Exists(Path.Combine(current.FullName, "ExcelCSIToolBox.sln")))
                    {
                        return current.FullName;
                    }

                    current = current.Parent;
                }

                throw new DirectoryNotFoundException("Could not locate repository root from " + AppDomain.CurrentDomain.BaseDirectory);
            }
        }
    }
}
