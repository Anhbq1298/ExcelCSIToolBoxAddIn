using System.Collections.Generic;
using ExcelCSIToolBox.Application.ToolCatalog.Contracts;

namespace ExcelCSIToolBox.Application.ToolCatalog.Registry
{
    public sealed class ToolSchemaRegistryBuilder
    {
        private readonly List<IToolSchemaModule> _modules = new List<IToolSchemaModule>();

        public ToolSchemaRegistryBuilder AddModule(IToolSchemaModule module)
        {
            if (module != null)
            {
                _modules.Add(module);
            }

            return this;
        }

        public ToolSchemaRegistry Build()
        {
            List<ToolSchema> schemas = new List<ToolSchema>();

            foreach (IToolSchemaModule module in _modules)
            {
                IEnumerable<ToolSchema> moduleSchemas = module.GetSchemas();

                if (moduleSchemas == null)
                {
                    continue;
                }

                schemas.AddRange(moduleSchemas);
            }

            return new ToolSchemaRegistry(schemas);
        }
    }
}
