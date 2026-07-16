using System.Collections.Generic;
using ExcelCSIToolBox.Application.Features.Loadings;
using ExcelCSIToolBox.Core.Contracts.CSI;
using FluentAssertions;
using Xunit;

namespace ExcelCSIToolBox.Tests.Application.Features.Loadings
{
    public class ShellUniformLoadSetSelectionPlannerTests
    {
        [Fact]
        public void NormalizeLoadSetNames_trims_removes_blanks_deduplicates_and_sorts()
        {
            IReadOnlyList<string> names = ShellUniformLoadSetSelectionPlanner.NormalizeLoadSetNames(
                new[] { " Office ", "", "roof", "OFFICE", null, "Lobby" });

            names.Should().Equal("Lobby", "Office", "roof");
        }

        [Fact]
        public void CreatePlan_selects_shells_assigned_to_any_selected_load_set_case_insensitively()
        {
            var assignments = new[]
            {
                Assignment("Story1", "F1", "101", "Office"),
                Assignment("Story1", "F2", "102", "Lobby"),
                Assignment("Story1", "F3", "103", "Storage")
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "office", " LOBBY " },
                new[] { "Office", "Lobby", "Storage" },
                assignments,
                assignment => null);

            plan.AreaObjectNames.Should().Equal("101", "102");
            plan.MatchedLoadSetNames.Should().Equal("Lobby", "Office");
            plan.MatchingAssignmentCount.Should().Be(2);
        }

        [Fact]
        public void CreatePlan_resolves_area_label_and_story_when_unique_name_is_missing()
        {
            var assignments = new[]
            {
                Assignment("Level 2", "F12", null, "Office")
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office" },
                new[] { "Office" },
                assignments,
                assignment => assignment.Story == "Level 2" && assignment.Label == "F12" ? "2201" : null);

            plan.AreaObjectNames.Should().Equal("2201");
            plan.UnresolvedAreaCount.Should().Be(0);
        }

        [Fact]
        public void CreatePlan_deduplicates_shell_objects_and_counts_duplicates()
        {
            var assignments = new[]
            {
                Assignment("Story1", "F1", "101", "Office"),
                Assignment("Story1", "F1", "101", "Lobby"),
                Assignment("Story1", "F2", "102", "Lobby")
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office", "Lobby" },
                new[] { "Office", "Lobby" },
                assignments,
                assignment => null);

            plan.AreaObjectNames.Should().Equal("101", "102");
            plan.DuplicateShellCount.Should().Be(1);
        }

        [Fact]
        public void CreatePlan_filters_matching_shells_to_selected_story()
        {
            var assignments = new[]
            {
                Assignment("Story1", "F1", "101", "Office"),
                Assignment("Story2", "F2", "102", "Office"),
                Assignment("Story2", "F3", "103", "Lobby")
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office" },
                new[] { "Office", "Lobby" },
                assignments,
                assignment => null,
                " story2 ");

            plan.SelectedStoryName.Should().Be("story2");
            plan.AreaObjectNames.Should().Equal("102");
            plan.MatchedLoadSetNames.Should().Equal("Office");
            plan.MatchingAssignmentCount.Should().Be(1);
        }

        [Fact]
        public void CreatePlan_filters_matching_shells_to_multiple_selected_stories()
        {
            var assignments = new[]
            {
                Assignment("Story1", "F1", "101", "Office"),
                Assignment("Story2", "F2", "102", "Office"),
                Assignment("Story3", "F3", "103", "Office"),
                Assignment("Story4", "F4", "104", "Office")
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office" },
                new[] { "Office" },
                assignments,
                assignment => null,
                new[] { "story1", " Story3 " });

            plan.SelectedStoryNames.Should().Equal("story1", "Story3");
            plan.AreaObjectNames.Should().Equal("101", "103");
            plan.MatchingAssignmentCount.Should().Be(2);
        }

        [Fact]
        public void CreatePlan_reports_unknown_selected_load_sets()
        {
            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Missing", "Office" },
                new[] { "Office" },
                new[] { Assignment("Story1", "F1", "101", "Office") },
                assignment => null);

            plan.UnknownLoadSetNames.Should().Equal("Missing");
            plan.AreaObjectNames.Should().Equal("101");
        }

        [Fact]
        public void CreatePlan_handles_no_matching_assignments()
        {
            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office" },
                new[] { "Office" },
                new[] { Assignment("Story1", "F1", "101", "Lobby") },
                assignment => null);

            plan.AreaObjectNames.Should().BeEmpty();
            plan.MatchingAssignmentCount.Should().Be(0);
        }

        [Fact]
        public void CreatePlan_ignores_malformed_rows_and_counts_unresolved_matching_areas()
        {
            var assignments = new[]
            {
                Assignment("Story1", "F1", null, "Office"),
                Assignment("Story1", "F2", "102", null),
                null
            };

            ShellUniformLoadSetSelectionPlan plan = ShellUniformLoadSetSelectionPlanner.CreatePlan(
                new[] { "Office" },
                new[] { "Office" },
                assignments,
                assignment => null);

            plan.AreaObjectNames.Should().BeEmpty();
            plan.MatchingAssignmentCount.Should().Be(1);
            plan.UnresolvedAreaCount.Should().Be(1);
            plan.UnresolvedAreaReferences.Should().Equal("Story1/F1");
        }

        private static ShellUniformLoadSetAreaAssignmentDto Assignment(
            string story,
            string label,
            string uniqueName,
            string loadSetName)
        {
            return new ShellUniformLoadSetAreaAssignmentDto
            {
                Story = story,
                Label = label,
                UniqueName = uniqueName,
                LoadSetName = loadSetName
            };
        }
    }
}
