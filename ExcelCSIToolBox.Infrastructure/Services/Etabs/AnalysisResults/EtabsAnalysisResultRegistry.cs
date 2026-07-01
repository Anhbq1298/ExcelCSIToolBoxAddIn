using System;
using System.Collections.ObjectModel;
using ExcelCSIToolBox.Core.Models.AnalysisResults;

namespace ExcelCSIToolBox.Infrastructure.Services.Etabs.AnalysisResults
{
    public static class EtabsAnalysisResultRegistry
    {
        public static ObservableCollection<AnalysisResultGroup> CreateDefaultGroups()
        {
            ObservableCollection<AnalysisResultGroup> groups = new ObservableCollection<AnalysisResultGroup>();

            AnalysisResultGroup displacements = new AnalysisResultGroup("Displacements");
            displacements.Items.Add(new AnalysisResultItem("Joint Displacements", "JOINT_DISPLACEMENTS", "Displacements", "Joint Displacements"));
            displacements.Items.Add(new AnalysisResultItem("Joint Displacements - Absolute", "JOINT_DISPLACEMENTS_ABSOLUTE", "Displacements", "Joint Displacements - Absolute"));
            displacements.Items.Add(new AnalysisResultItem("Joint Drifts", "JOINT_DRIFTS", "Displacements", "Joint Drifts"));
            displacements.Items.Add(new AnalysisResultItem("Diaphragm Center Of Mass Displacements", "DIAPHRAGM_CENTER_OF_MASS_DISPLACEMENTS", "Displacements", "Diaphragm Center Of Mass Displacements"));
            displacements.Items.Add(new AnalysisResultItem("Diaphragm Max Over Avg Drifts", "DIAPHRAGM_MAX_OVER_AVG_DRIFTS", "Displacements", "Diaphragm Max Over Avg Drifts"));
            displacements.Items.Add(new AnalysisResultItem("Story Drifts", "STORY_DRIFTS", "Displacements", "Story Drifts"));
            displacements.Items.Add(new AnalysisResultItem("Story Max Over Avg Displacements", "STORY_MAX_OVER_AVG_DISPLACEMENTS", "Displacements", "Story Max Over Avg Displacements"));
            displacements.Items.Add(new AnalysisResultItem("Story Max Over Avg Drifts", "STORY_MAX_OVER_AVG_DRIFTS", "Displacements", "Story Max Over Avg Drifts"));

            AnalysisResultGroup reactions = new AnalysisResultGroup("Reactions");
            reactions.Items.Add(new AnalysisResultItem("Joint Reactions", "JOINT_REACTIONS", "Reactions", "Joint Reactions"));
            reactions.Items.Add(new AnalysisResultItem("Joint Design Reactions", "JOINT_DESIGN_REACTIONS", "Reactions", "Joint Design Reactions"));

            AnalysisResultGroup velocityAndAcceleration = new AnalysisResultGroup("Velocity and Acceleration");
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Joint Velocities - Relative", "JOINT_VELOCITIES_RELATIVE", "Velocity and Acceleration", "Joint Velocities - Relative"));
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Joint Velocities - Absolute", "JOINT_VELOCITIES_ABSOLUTE", "Velocity and Acceleration", "Joint Velocities - Absolute"));
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Joint Accelerations - Relative", "JOINT_ACCELERATIONS_RELATIVE", "Velocity and Acceleration", "Joint Accelerations - Relative"));
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Joint Accelerations - Absolute", "JOINT_ACCELERATIONS_ABSOLUTE", "Velocity and Acceleration", "Joint Accelerations - Absolute"));
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Diaphragm Accelerations", "DIAPHRAGM_ACCELERATIONS", "Velocity and Acceleration", "Diaphragm Accelerations"));
            velocityAndAcceleration.Items.Add(new AnalysisResultItem("Story Accelerations", "STORY_ACCELERATIONS", "Velocity and Acceleration", "Story Accelerations"));

            AnalysisResultGroup jointMasses = new AnalysisResultGroup("Assembled Joint Masses");
            jointMasses.Items.Add(new AnalysisResultItem("Assembled Joint Masses", "JOINT_MASSES", "Joint Masses", "Assembled Joint Masses"));

            AnalysisResultGroup frameOutput = new AnalysisResultGroup("Frame Output");
            frameOutput.Items.Add(new AnalysisResultItem("Element Forces - Columns", "FRAME_OUTPUT_COLUMNS", "Frame Output", "Element Forces - Columns"));
            frameOutput.Items.Add(new AnalysisResultItem("Element Forces - Beams", "FRAME_OUTPUT_BEAMS", "Frame Output", "Element Forces - Beams"));
            frameOutput.Items.Add(new AnalysisResultItem("Element Forces - Braces", "FRAME_OUTPUT_BRACES", "Frame Output", "Element Forces - Braces"));
            frameOutput.Items.Add(new AnalysisResultItem("Element Joint Forces - Frame", "FRAME_OUTPUT_JOINT_FORCES", "Frame Output", "Element Joint Forces - Frame"));

            AnalysisResultGroup areaOutput = new AnalysisResultGroup("Area Output");
            areaOutput.Items.Add(new AnalysisResultItem("Element Forces - Area Shells", "AREA_OUTPUT_FORCES", "Area Output", "Element Forces - Area Shells"));
            areaOutput.Items.Add(new AnalysisResultItem("Element Stresses - Area Shells", "AREA_OUTPUT_STRESSES", "Area Output", "Element Stresses - Area Shells"));
            areaOutput.Items.Add(new AnalysisResultItem("Element Strains - Area Shells", "AREA_OUTPUT_STRAINS", "Area Output", "Element Strains - Area Shells"));
            areaOutput.Items.Add(new AnalysisResultItem("Element Joint Forces - Shells", "AREA_OUTPUT_JOINT_FORCES", "Area Output", "Element Joint Forces - Shells"));

            AnalysisResultGroup wallOutput = new AnalysisResultGroup("Wall Output");
            wallOutput.Items.Add(new AnalysisResultItem("Pier Forces", "WALL_OUTPUT_PIER_FORCES", "Wall Output", "Pier Forces"));

            AnalysisResultGroup objectsAndElements = new AnalysisResultGroup("Objects and Elements");
            objectsAndElements.Items.Add(new AnalysisResultItem("Objects and Elements - Joints", "OBJECTS_AND_ELEMENTS_JOINTS", "Objects and Elements", "Objects and Elements - Joints"));
            objectsAndElements.Items.Add(new AnalysisResultItem("Objects and Elements - Frames", "OBJECTS_AND_ELEMENTS_FRAMES", "Objects and Elements", "Objects and Elements - Frames"));
            objectsAndElements.Items.Add(new AnalysisResultItem("Objects and Elements - Areas", "OBJECTS_AND_ELEMENTS_AREAS", "Objects and Elements", "Objects and Elements - Areas"));

            AnalysisResultGroup baseReactions = new AnalysisResultGroup("Base Reactions");
            baseReactions.Items.Add(new AnalysisResultItem("Base Reactions", "BASE_REACTIONS", "Base Reactions", "Base Reactions"));

            AnalysisResultGroup modalInformation = new AnalysisResultGroup("Modal Information");
            modalInformation.Items.Add(new AnalysisResultItem("Modal Periods And Frequencies", "MODAL_PERIODS_AND_FREQUENCIES", "Modal Information", "Modal Periods And Frequencies"));
            modalInformation.Items.Add(new AnalysisResultItem("Modal Participating Mass Ratios", "MODAL_INFORMATION", "Modal Information", "Modal Participating Mass Ratios"));
            modalInformation.Items.Add(new AnalysisResultItem("Modal Load Participation Ratios", "MODAL_LOAD_PARTICIPATION_RATIOS", "Modal Information", "Modal Load Participation Ratios"));
            modalInformation.Items.Add(new AnalysisResultItem("Modal Participation Factors", "MODAL_PARTICIPATION_FACTORS", "Modal Information", "Modal Participation Factors"));
            modalInformation.Items.Add(new AnalysisResultItem("Modal Direction Factors", "MODAL_DIRECTION_FACTORS", "Modal Information", "Modal Direction Factors"));
            modalInformation.Items.Add(new AnalysisResultItem("Response Spectrum Modal Info", "RESPONSE_SPECTRUM_MODAL_INFO", "Modal Information", "Response Spectrum Modal Info"));

            AnalysisResultGroup massData = new AnalysisResultGroup("Mass Data");
            massData.Items.Add(new AnalysisResultItem("Mass Summary by Story", "MASS_DATA", "Mass Data", "Mass Summary by Story"));
            massData.Items.Add(new AnalysisResultItem("Mass Summary by Diaphragm", "MASS_SUMMARY_BY_DIAPHRAGM", "Mass Data", "Mass Summary by Diaphragm"));
            massData.Items.Add(new AnalysisResultItem("Mass Summary by Group", "MASS_SUMMARY_BY_GROUP", "Mass Data", "Mass Summary by Group"));

            AnalysisResultGroup otherOutputItems = new AnalysisResultGroup("Other Output Items");
            otherOutputItems.Items.Add(new AnalysisResultItem("Centers Of Mass And Rigidity", "CENTERS_OF_MASS_AND_RIGIDITY", "Other Output Items", "Centers Of Mass And Rigidity"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Story Forces", "STORY_FORCES", "Other Output Items", "Story Forces"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Diaphragm Forces", "DIAPHRAGM_FORCES", "Other Output Items", "Diaphragm Forces"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Story Stiffness", "STORY_STIFFNESS", "Other Output Items", "Story Stiffness"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Shear Gravity Ratios", "SHEAR_GRAVITY_RATIOS", "Other Output Items", "Shear Gravity Ratios"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Stiffness Gravity Ratios", "STIFFNESS_GRAVITY_RATIOS", "Other Output Items", "Stiffness Gravity Ratios"));
            otherOutputItems.Items.Add(new AnalysisResultItem("Tributary Area and LLRF", "TRIBUTARY_AREA_AND_LLRF", "Other Output Items", "Tributary Area and LLRF"));

            AnalysisResultGroup connectivity = new AnalysisResultGroup("Etabs Object Connectivity");
            connectivity.Items.Add(new AnalysisResultItem("Point Object Connectivity", "POINT_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Point Object Connectivity"));
            connectivity.Items.Add(new AnalysisResultItem("Beam Object Connectivity", "BEAM_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Beam Object Connectivity"));
            connectivity.Items.Add(new AnalysisResultItem("Column Object Connectivity", "COLUMN_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Column Object Connectivity"));
            connectivity.Items.Add(new AnalysisResultItem("Brace Object Connectivity", "BRACE_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Brace Object Connectivity"));
            connectivity.Items.Add(new AnalysisResultItem("Floor Object Connectivity", "FLOOR_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Floor Object Connectivity"));
            connectivity.Items.Add(new AnalysisResultItem("Wall Object Connectivity", "WALL_OBJECT_CONNECTIVITY", "Etabs Object Connectivity", "Wall Object Connectivity"));

            AnalysisResultGroup projectInformation = new AnalysisResultGroup("Project Information");
            projectInformation.Items.Add(new AnalysisResultItem("Project Information", "PROJECT_INFORMATION", "Project Information", "Project Information"));

            AnalysisResultGroup materialList = new AnalysisResultGroup("Material List");
            materialList.Items.Add(new AnalysisResultItem("Material List by Object Type", "MATERIAL_LIST_BY_OBJECT_TYPE", "Material List", "Material List by Object Type"));
            materialList.Items.Add(new AnalysisResultItem("Material List by Section Property", "MATERIAL_LIST_BY_SECTION_PROPERTY", "Material List", "Material List by Section Property"));
            materialList.Items.Add(new AnalysisResultItem("Material List by Story", "MATERIAL_LIST_BY_STORY", "Material List", "Material List by Story"));

            groups.Add(displacements);
            groups.Add(reactions);
            groups.Add(velocityAndAcceleration);
            groups.Add(jointMasses);
            groups.Add(frameOutput);
            groups.Add(areaOutput);
            groups.Add(wallOutput);
            groups.Add(objectsAndElements);
            groups.Add(baseReactions);
            groups.Add(modalInformation);
            groups.Add(massData);
            groups.Add(otherOutputItems);
            groups.Add(connectivity);
            groups.Add(projectInformation);
            groups.Add(materialList);

            return groups;
        }

        public static AnalysisResultGroup CreateGroupForNavigation(string groupName)
        {
            string normalizedGroupName = NormalizeNavigationGroupName(groupName);
            foreach (AnalysisResultGroup group in CreateDefaultGroups())
            {
                if (string.Equals(group.Name, normalizedGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    return group;
                }
            }

            return new AnalysisResultGroup(normalizedGroupName);
        }

        public static ObservableCollection<string> GetSupportedKeysForGenericTableExport()
        {
            ObservableCollection<string> keys = new ObservableCollection<string>();
            foreach (AnalysisResultGroup group in CreateDefaultGroups())
            {
                foreach (AnalysisResultItem item in group.Items)
                {
                    if (!string.Equals(item.Key, "JOINT_DISPLACEMENTS", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(item.Key, "BASE_REACTIONS", StringComparison.OrdinalIgnoreCase))
                    {
                        keys.Add(item.Key);
                    }
                }
            }

            return keys;
        }

        private static string NormalizeNavigationGroupName(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
            {
                return "Base Reactions";
            }

            switch (groupName)
            {
                case "Joint Output":
                case "Joint Displacements":
                case "Joint Displacements - Absolute":
                case "Joint Drifts":
                case "Diaphragm Center Of Mass Displacements":
                case "Diaphragm Max Over Avg Drifts":
                case "Story Drifts":
                case "Story Max Over Avg Displacements":
                case "Story Max Over Avg Drifts":
                    return "Displacements";
                case "Joint Reactions":
                case "Joint Design Reactions":
                    return "Reactions";
                case "Joint Velocities - Relative":
                case "Joint Velocities - Absolute":
                case "Joint Accelerations - Relative":
                case "Joint Accelerations - Absolute":
                case "Diaphragm Accelerations":
                case "Story Accelerations":
                    return "Velocity and Acceleration";
                case "Joint Masses":
                    return "Assembled Joint Masses";
                case "Mass Summary by Story":
                case "Mass Summary by Diaphragm":
                case "Mass Summary by Group":
                    return "Mass Data";
                case "Point Object Connectivity":
                case "Beam Object Connectivity":
                case "Column Object Connectivity":
                case "Brace Object Connectivity":
                case "Floor Object Connectivity":
                case "Wall Object Connectivity":
                    return "Etabs Object Connectivity";
                default:
                    return groupName;
            }
        }
    }
}
