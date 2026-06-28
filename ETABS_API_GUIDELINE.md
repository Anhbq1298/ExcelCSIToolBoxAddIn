# ETABS API Integration Guideline (Unique Names vs. Labels & Database Tables)

This document outlines key technical behaviors of the ETABS Open Application Programming Interface (OAPI) to prevent common mistakes when implementing selection-only filtering, object mappings, or data extraction workflows.

---

## 1. Selection API & Identifiers

### Unique Names vs. Labels
- **Unique Name**: A unique numeric string identifier generated internally by ETABS (e.g., `"1"`, `"2"`, `"3"`) to represent joints, frames, and area objects.
- **Label**: A user-friendly name (e.g., `"C1"`, `"B1"`, `"W1"`) assigned to elements, often scoped by story.
- **Critical OAPI Behavior**: The selection method `sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames)` **always returns the Unique Names** of selected objects in the `objectNames` array.

### Selection Resolution Guideline
When querying the active selection to filter objects or database tables, you must resolve **both the Unique Name and the corresponding Label** for each selected element.
To do this:
1. Query the selected unique names.
2. For each unique name, query its label using:
   - Points: `sapModel.PointObj.GetLabelFromName(uniqueName, ref label, ref story)`
   - Frames: `sapModel.FrameObj.GetLabelFromName(uniqueName, ref label, ref story)`
   - Areas: `sapModel.AreaObj.GetLabelFromName(uniqueName, ref label, ref story)`
3. Store both the **Unique Name** and the **Label** in the selection matching set. This guarantees a match regardless of how the database table references the structural element.

---

## 2. Database Table Headers (`GetTableForDisplayArray`)

When querying database tables via `sapModel.DatabaseTables.GetTableForDisplayArray`, the field keys (`fieldKeyList`) returned vary depending on the table and element type. 
- Standard tables **rarely** contain a `"UniqueName"` or `"Unique Name"` column.
- Instead, elements are identified by their user-friendly labels in element-specific columns.

### Element Column Names Catalog
Always map the element column index using a case-insensitive search across these candidate names:

| Category | Table Type (Examples) | Candidate Field Keys |
| :--- | :--- | :--- |
| **Joint / Point** | Joint Displacements, Reactions, Drifts | `"Unique Name"`, `"UniqueName"`, `"Joint"`, `"Joint Name"`, `"JointName"`, `"Point"`, `"Point Name"`, `"PointName"`, `"Label"`, `"Label Name"`, `"LabelName"` |
| **Frame** | Element Forces - Columns / Beams / Braces | `"Unique Name"`, `"UniqueName"`, `"Frame"`, `"Frame Name"`, `"FrameName"`, **`"Column"`**, **`"Beam"`**, **`"Brace"`**, `"Element"`, `"Element Name"`, `"ElementName"`, `"Label"` |
| **Area / Shell** | Element Forces / Stresses - Area Shells | `"Unique Name"`, `"UniqueName"`, `"Area"`, `"Area Name"`, `"AreaName"`, **`"Shell"`**, **`"Shell Name"`**, **`"ShellName"`**, `"Element"`, `"Element Name"`, `"ElementName"`, `"Label"` |
| **Wall / Pier** | Pier Forces, Spandrel Forces | `"Pier"`, `"Pier Name"`, `"PierName"`, `"Spandrel"`, `"Spandrel Name"`, `"SpandrelName"`, `"Label"` |

---

## 3. Implementation Example (C#)

Always structure selection filtering to match both unique names and user-facing labels:

```csharp
public HashSet<string> GetActiveFramesSelection(ETABSv1.cSapModel sapModel)
{
    var selectedFrames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    try
    {
        int numberItems = 0;
        int[] objectTypes = null;
        string[] objectNames = null;
        int ret = sapModel.SelectObj.GetSelected(ref numberItems, ref objectTypes, ref objectNames);
        if (ret == 0 && numberItems > 0 && objectTypes != null && objectNames != null)
        {
            for (int i = 0; i < numberItems; i++)
            {
                if (objectTypes[i] == FrameObjectType && !string.IsNullOrWhiteSpace(objectNames[i]))
                {
                    string uniqueName = objectNames[i].Trim();
                    selectedFrames.Add(uniqueName); // Add Unique Name

                    string label = string.Empty;
                    string story = string.Empty;
                    // Add Label
                    if (sapModel.FrameObj.GetLabelFromName(uniqueName, ref label, ref story) == 0 && !string.IsNullOrWhiteSpace(label))
                    {
                        selectedFrames.Add(label.Trim());
                    }
                }
            }
        }
    }
    catch
    {
        // Fallback or ignore
    }
    return selectedFrames;
}
```
