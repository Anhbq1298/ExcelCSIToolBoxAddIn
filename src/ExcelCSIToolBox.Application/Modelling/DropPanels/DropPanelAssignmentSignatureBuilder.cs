using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelCSIToolBox.Application.Modelling.DropPanels
{
    public sealed class DropPanelAssignmentSignatureBuilder
    {
        public string Build(string resultingSectionProperty, DropPanelAreaAssignmentBackup assignment)
        {
            if (assignment == null)
            {
                throw new ArgumentNullException(nameof(assignment));
            }

            StringBuilder value = new StringBuilder();
            Append(value, "Property", resultingSectionProperty);
            Append(value, "Axis", Format(assignment.LocalAxisAngle));
            Append(value, "Local3", Vector(assignment.Local3Direction));
            Append(value, "Diaphragm", assignment.Diaphragm);
            Append(value, "Pier", assignment.PierLabel);
            Append(value, "Spandrel", assignment.SpandrelLabel);
            Append(value, "Groups", Join(assignment.Groups));
            Append(value, "Modifiers", JoinNumbers(assignment.Modifiers));
            Append(value, "LoadSets", Join(assignment.ShellUniformLoadSetNames));

            IEnumerable<string> directLoads = assignment.DirectAreaLoads
                .Select(load => Normalize(load.LoadPattern) + "|" +
                                Normalize(load.LoadType) + "|" +
                                load.Direction.ToString(CultureInfo.InvariantCulture) + "|" +
                                Normalize(load.CoordinateSystem) + "|" +
                                Format(load.Value) + "|" +
                                load.ReplaceExistingAssignments.ToString(CultureInfo.InvariantCulture))
                .OrderBy(item => item, StringComparer.Ordinal);
            Append(value, "DirectLoads", string.Join(";", directLoads));

            if (assignment.MeshAssignment != null)
            {
                List<string> meshRecords = new List<string>();
                foreach (DropPanelTableRecord record in assignment.MeshAssignment.Records)
                {
                    IEnumerable<string> fields = record.Values
                        .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(pair => Normalize(pair.Key) + "=" + Normalize(pair.Value));
                    meshRecords.Add(string.Join("|", fields));
                }

                meshRecords.Sort(StringComparer.Ordinal);
                Append(value, "Mesh", Normalize(assignment.MeshAssignment.TableKey) + ":" + string.Join(";", meshRecords));
            }

            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value.ToString()));
                StringBuilder result = new StringBuilder(hash.Length * 2);
                for (int index = 0; index < hash.Length; index++)
                {
                    result.Append(hash[index].ToString("x2", CultureInfo.InvariantCulture));
                }

                return result.ToString();
            }
        }

        private static void Append(StringBuilder value, string name, string fieldValue)
        {
            value.Append(name).Append('=').Append(Normalize(fieldValue)).Append('\n');
        }

        private static string Join(IEnumerable<string> values)
        {
            return string.Join(";", (values ?? new string[0])
                .Select(Normalize)
                .OrderBy(item => item, StringComparer.Ordinal));
        }

        private static string JoinNumbers(IEnumerable<double> values)
        {
            return string.Join(";", (values ?? new double[0]).Select(Format));
        }

        private static string Vector(DropPanelVector3D vector)
        {
            return vector == null
                ? string.Empty
                : Format(vector.X) + "," + Format(vector.Y) + "," + Format(vector.Z);
        }

        private static string Format(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }
    }
}
