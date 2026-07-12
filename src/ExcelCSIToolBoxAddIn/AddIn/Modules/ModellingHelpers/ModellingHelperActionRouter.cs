using System;
using System.Collections.Generic;

namespace ExcelCSIToolBoxAddIn.AddIn.Modules.ModellingHelpers
{
    public class ModellingHelperActionRouter
    {
        private readonly Dictionary<string, Action> _actions;

        public ModellingHelperActionRouter()
        {
            _actions = new Dictionary<string, Action>(StringComparer.OrdinalIgnoreCase);
        }

        public ModellingHelperActionRouter Register(string key, Action action)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new ArgumentException("Modelling helper action key is required.", nameof(key));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            _actions[key] = action;
            return this;
        }

        public void Execute(string key)
        {
            Action action;
            if (!_actions.TryGetValue(key, out action))
            {
                throw new InvalidOperationException("No modelling helper action registered for key: " + key);
            }

            action();
        }
    }
}
