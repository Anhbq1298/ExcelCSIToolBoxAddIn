using System;
using System.IO;
using System.Xml.Serialization;
using ExcelCSIToolBox.Application.Modelling.DropPanels;
using ExcelCSIToolBoxAddIn.AddIn.Diagnostics;

namespace ExcelCSIToolBoxAddIn.UI.ViewModels
{
    public sealed class DropPanelSettingsStore
    {
        public DropPanelOptions Load()
        {
            string xml = Properties.Settings.Default.DropPanelSettingsXml;
            if (string.IsNullOrWhiteSpace(xml))
            {
                return new DropPanelOptions();
            }

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(DropPanelOptions));
                using (StringReader reader = new StringReader(xml))
                {
                    return serializer.Deserialize(reader) as DropPanelOptions ?? new DropPanelOptions();
                }
            }
            catch (Exception ex)
            {
                AddInDiagnostics.LogException("Load Drop Panel settings", ex);
                return new DropPanelOptions();
            }
        }

        public void Save(DropPanelOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            XmlSerializer serializer = new XmlSerializer(typeof(DropPanelOptions));
            using (StringWriter writer = new StringWriter(System.Globalization.CultureInfo.InvariantCulture))
            {
                serializer.Serialize(writer, options);
                Properties.Settings.Default.DropPanelSettingsXml = writer.ToString();
                Properties.Settings.Default.Save();
            }
        }
    }
}
