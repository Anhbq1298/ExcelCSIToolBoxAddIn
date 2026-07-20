# SAP2000 Cable Pretension Setter

Standalone WinForms utility for assigning cable pretension values to cable objects in the active SAP2000 model.

## Scope

- Attach to a running SAP2000 instance.
- Read all available cable objects from the active model.
- Display cable name, I/J joints, current cable definition and current tension values.
- Select individual cables or all cables.
- Assign one of these SAP2000 cable definition types:
  - Tension at I-End (`CableType = 3`)
  - Tension at J-End (`CableType = 4`)
  - Horizontal tension component (`CableType = 5`)
- Preserve the cable's current number of segments, added weight, projected load, deformed-geometry option and frame-modelling option.
- Verify each assignment by reading the cable definition back from SAP2000.
- Log success or failure for every cable.

## Safety behavior

- The app blocks writes when the SAP2000 model is locked.
- Values are absolute replacements in the active model's current force unit.
- The app does not run analysis and does not save the model automatically.
- A confirmation dialog is shown before any write operation.

## API implementation

The tool uses COM late binding to avoid a hard dependency on a specific SAP2000 interop assembly version. The write operation is performed with:

```text
SapModel.CableObj.SetCableData(
    Name,
    CableType,
    NumSegs,
    Weight,
    ProjectedLoad,
    Value,
    UseDeformedGeom,
    ModelUsingFrames)
```

The existing non-pretension cable parameters are first read with `GetCableData` and then passed back unchanged.

## Build

Open `Sap2000CablePretensionSetter.sln` in Visual Studio 2022 and build `Release | Any CPU`.

Requirements:

- Windows
- .NET Framework 4.8
- SAP2000 installed and running when using the app

The executable is generated under:

```text
bin\Release\net48\Sap2000CablePretensionSetter.exe
```
