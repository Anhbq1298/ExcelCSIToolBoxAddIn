using System;
using System.IO;

namespace ExcelCSIToolBox.RefBuilder.Chm
{
    public sealed class ChmExtractor : IChmExtractor
    {
        public void Extract(string refRoot)
        {
            if (!Directory.Exists(refRoot))
            {
                Directory.CreateDirectory(refRoot);
            }

            string markerPath = Path.Combine(refRoot, "extraction-note.txt");
            File.WriteAllText(
                markerPath,
                "CHM extraction is a development-time step. This fallback records discovered CHM files; replace IChmExtractor with an hh.exe/7zip extractor when needed." + Environment.NewLine);
        }
    }
}
