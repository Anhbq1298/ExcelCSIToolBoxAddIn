using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml.Linq;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public static class OutputTablePopupProfileProvider
    {
        private const string ProfileFileName = "OutputTablePopupProfiles.xml";
        private static readonly object SyncRoot = new object();
        private static IDictionary<string, OutputTablePopupProfile> _profiles;

        public static OutputTablePopupProfile GetProfile(string key)
        {
            EnsureLoaded();

            OutputTablePopupProfile profile;
            if (!string.IsNullOrWhiteSpace(key) && _profiles.TryGetValue(key, out profile))
            {
                return profile;
            }

            if (_profiles.TryGetValue("ForceOutput", out profile))
            {
                return profile;
            }

            return new OutputTablePopupProfile();
        }

        private static void EnsureLoaded()
        {
            if (_profiles != null)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (_profiles != null)
                {
                    return;
                }

                _profiles = LoadProfiles();
            }
        }

        private static IDictionary<string, OutputTablePopupProfile> LoadProfiles()
        {
            var profiles = CreateDefaultProfiles();
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI", "Config", ProfileFileName);
            if (!File.Exists(path))
            {
                path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ProfileFileName);
            }

            if (!File.Exists(path))
            {
                return profiles;
            }

            try
            {
                XDocument document = XDocument.Load(path);
                foreach (XElement element in document.Root == null
                    ? new XElement[0]
                    : document.Root.Elements("Profile"))
                {
                    OutputTablePopupProfile profile = ReadProfile(element);
                    if (!string.IsNullOrWhiteSpace(profile.Key))
                    {
                        profiles[profile.Key] = profile;
                    }
                }
            }
            catch (Exception ex)
            {
                AnalysisExportDiagnostics.Log("Failed to load output table popup profiles: " + ex.Message);
                // Popup profiles are optional configuration; fall back to built-in defaults.
            }

            return profiles;
        }

        private static IDictionary<string, OutputTablePopupProfile> CreateDefaultProfiles()
        {
            return new Dictionary<string, OutputTablePopupProfile>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "ModalInformation",
                    new OutputTablePopupProfile
                    {
                        Key = "ModalInformation",
                        CaseSelectionMode = OutputCaseSelectionMode.ModalCaseOnly,
                        CaseSelectorTitle = "Modal Case",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No modal result records found.",
                        WorksheetNamePrefix = "Modal"
                    }
                },
                {
                    "ResponseSpectrumModalInfo",
                    new OutputTablePopupProfile
                    {
                        Key = "ResponseSpectrumModalInfo",
                        CaseSelectionMode = OutputCaseSelectionMode.ResponseSpectrumCaseOnly,
                        CaseSelectorTitle = "Response Spectrum Case",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No response spectrum modal records found.",
                        WorksheetNamePrefix = "ResponseSpectrumModal"
                    }
                },
                {
                    "ForceOutput",
                    new OutputTablePopupProfile()
                },
                {
                    "ProjectInformation",
                    new OutputTablePopupProfile
                    {
                        Key = "ProjectInformation",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No project information records found.",
                        WorksheetNamePrefix = "Project Info"
                    }
                },
                {
                    "MaterialList",
                    new OutputTablePopupProfile
                    {
                        Key = "MaterialList",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = true,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No material list records found.",
                        WorksheetNamePrefix = "Material List"
                    }
                },
                {
                    "MassData",
                    new OutputTablePopupProfile
                    {
                        Key = "MassData",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = true,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No mass data records found.",
                        WorksheetNamePrefix = "Mass Data"
                    }
                },
                {
                    "StoryForces",
                    new OutputTablePopupProfile
                    {
                        Key = "StoryForces",
                        CaseSelectionMode = OutputCaseSelectionMode.AllCasesAndCombos,
                        CaseSelectorTitle = "Load Case / Combination",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = true,
                        ShowComboSelector = true,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No story force records found.",
                        WorksheetNamePrefix = "Story Forces"
                    }
                },
                {
                    "OtherOutputWithUnit",
                    new OutputTablePopupProfile
                    {
                        Key = "OtherOutputWithUnit",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = true,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No records found.",
                        WorksheetNamePrefix = "Other Output"
                    }
                },
                {
                    "OtherOutputRatioOnly",
                    new OutputTablePopupProfile
                    {
                        Key = "OtherOutputRatioOnly",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No ratio records found.",
                        WorksheetNamePrefix = "Ratio"
                    }
                },
                {
                    "DiaphragmForces",
                    new OutputTablePopupProfile
                    {
                        Key = "DiaphragmForces",
                        CaseSelectionMode = OutputCaseSelectionMode.AllCasesAndCombos,
                        CaseSelectorTitle = "Load Case / Combination",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = true,
                        ShowComboSelector = true,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No diaphragm force records found.",
                        WorksheetNamePrefix = "Diaphragm Forces"
                    }
                },
                {
                    "SeismicWindOrRSOnlyWithUnit",
                    new OutputTablePopupProfile
                    {
                        Key = "SeismicWindOrRSOnlyWithUnit",
                        CaseSelectionMode = OutputCaseSelectionMode.SeismicWindOrResponseSpectrumCasesOnly,
                        CaseSelectorTitle = "Load Case",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = true,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No story stiffness records found.",
                        WorksheetNamePrefix = "Story Stiffness"
                    }
                },
                {
                    "SeismicWindOrRSOnlyRatio",
                    new OutputTablePopupProfile
                    {
                        Key = "SeismicWindOrRSOnlyRatio",
                        CaseSelectionMode = OutputCaseSelectionMode.SeismicWindOrResponseSpectrumCasesOnly,
                        CaseSelectorTitle = "Load Case",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No ratio records found.",
                        WorksheetNamePrefix = "Ratio"
                    }
                },
                {
                    "JointOutput",
                    new OutputTablePopupProfile
                    {
                        Key = "JointOutput",
                        CaseSelectionMode = OutputCaseSelectionMode.AllCasesAndCombos,
                        CaseSelectorTitle = "Load Case / Combination",
                        AllowMultipleCases = true,
                        ShowCaseComboSelector = true,
                        ShowUnitSelector = true,
                        ShowComboSelector = true,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No joint output records found.",
                        WorksheetNamePrefix = "Joint Output"
                    }
                },
                {
                    "ObjectsAndElements",
                    new OutputTablePopupProfile
                    {
                        Key = "ObjectsAndElements",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = true,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = true,
                        EmptyDataMessage = "No objects/elements found.",
                        WorksheetNamePrefix = "Objects and Elements"
                    }
                },
                {
                    "EtabsObjectConnectivity",
                    new OutputTablePopupProfile
                    {
                        Key = "EtabsObjectConnectivity",
                        CaseSelectionMode = OutputCaseSelectionMode.None,
                        CaseSelectorTitle = string.Empty,
                        AllowMultipleCases = false,
                        ShowCaseComboSelector = false,
                        ShowUnitSelector = false,
                        ShowComboSelector = false,
                        DefaultToCurrentEtabsUnit = false,
                        EmptyDataMessage = "No object connectivity records found.",
                        WorksheetNamePrefix = "Object Connectivity"
                    }
                }
            };
        }

        private static OutputTablePopupProfile ReadProfile(XElement element)
        {
            var profile = new OutputTablePopupProfile
            {
                Key = ReadAttribute(element, "key", "ForceOutput"),
                CaseSelectionMode = ReadEnum(element, "CaseSelectionMode", OutputCaseSelectionMode.AllCasesAndCombos),
                CaseSelectorTitle = ReadString(element, "CaseSelectorTitle", "Load Case / Combination"),
                AllowMultipleCases = ReadBool(element, "AllowMultipleCases", true),
                ShowCaseComboSelector = ReadBool(element, "ShowCaseComboSelector", true),
                ShowUnitSelector = ReadBool(element, "ShowUnitSelector", true),
                ShowComboSelector = ReadBool(element, "ShowComboSelector", true),
                DefaultToCurrentEtabsUnit = ReadBool(element, "DefaultToCurrentEtabsUnit", true),
                EmptyDataMessage = ReadString(element, "EmptyDataMessage", "No records found."),
                WorksheetNamePrefix = ReadString(element, "WorksheetNamePrefix", "Output")
            };

            return profile;
        }

        private static string ReadAttribute(XElement element, string name, string defaultValue)
        {
            XAttribute attribute = element == null ? null : element.Attribute(name);
            return attribute == null || string.IsNullOrWhiteSpace(attribute.Value)
                ? defaultValue
                : attribute.Value.Trim();
        }

        private static string ReadString(XElement element, string name, string defaultValue)
        {
            XElement child = element == null ? null : element.Element(name);
            return child == null || string.IsNullOrWhiteSpace(child.Value)
                ? defaultValue
                : child.Value.Trim();
        }

        private static bool ReadBool(XElement element, string name, bool defaultValue)
        {
            string value = ReadString(element, name, string.Empty);
            bool result;
            return bool.TryParse(value, out result) ? result : defaultValue;
        }

        private static OutputCaseSelectionMode ReadEnum(XElement element, string name, OutputCaseSelectionMode defaultValue)
        {
            string value = ReadString(element, name, string.Empty);
            OutputCaseSelectionMode result;
            return Enum.TryParse(value, true, out result) ? result : defaultValue;
        }
    }
}
