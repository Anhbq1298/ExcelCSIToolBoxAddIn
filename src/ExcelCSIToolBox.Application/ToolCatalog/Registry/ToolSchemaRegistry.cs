using System;
using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry
{
    public sealed class ToolSchemaRegistry
    {
        private readonly List<ToolSchema> _schemas;

        public ToolSchemaRegistry(IEnumerable<ToolSchema> schemas)
        {
            _schemas = new List<ToolSchema>();
            if (schemas == null)
            {
                return;
            }

            foreach (ToolSchema schema in schemas)
            {
                if (schema != null && !string.IsNullOrWhiteSpace(schema.ToolName))
                {
                    _schemas.Add(schema);
                }
            }
        }

        public IReadOnlyList<ToolSchema> GetAll()
        {
            return new List<ToolSchema>(_schemas);
        }

        public ToolSchema FindByToolName(string toolName)
        {
            return FindByNameCore(toolName, false);
        }

        public ToolSchema FindByAlias(string alias)
        {
            return FindByNameCore(alias, true);
        }

        public IReadOnlyList<ToolSchema> FindByDomain(string domain)
        {
            var matches = new List<ToolSchema>();
            for (int i = 0; i < _schemas.Count; i++)
            {
                ToolSchema schema = _schemas[i];
                if (EqualsText(schema.Domain, domain))
                {
                    matches.Add(schema);
                }
            }

            return matches;
        }

        public IReadOnlyList<ToolSchema> FindByActionAndTarget(string action, string targetObject)
        {
            var matches = new List<ToolSchema>();
            for (int i = 0; i < _schemas.Count; i++)
            {
                ToolSchema schema = _schemas[i];
                if (EqualsText(schema.Action, action) &&
                    EqualsText(schema.TargetObject, targetObject))
                {
                    matches.Add(schema);
                }
            }

            return matches;
        }

        private ToolSchema FindByNameCore(string name, bool aliasOnly)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            for (int i = 0; i < _schemas.Count; i++)
            {
                ToolSchema schema = _schemas[i];
                if (!aliasOnly && EqualsText(schema.ToolName, name))
                {
                    return schema;
                }

                if (schema.Aliases == null)
                {
                    continue;
                }

                for (int aliasIndex = 0; aliasIndex < schema.Aliases.Count; aliasIndex++)
                {
                    if (EqualsText(schema.Aliases[aliasIndex], name))
                    {
                        return schema;
                    }
                }
            }

            return null;
        }

        private static bool EqualsText(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }
}
