# Helion NX TC Admin Tools

Administration tools for Helion's Siemens NX / Teamcenter environment.

## Contents

| Folder | Description |
|---|---|
| [MaterialEditor/](MaterialEditor/) | Excel-based editor for the NX MatML material library (`helion_material_library.xml`). Export to a spreadsheet, edit Name / Class / Subclass / `MassDensityConstant_110`, import back to XML with an automatic backup. See [MaterialEditor/README.md](MaterialEditor/README.md). |
| [MaterialEditor/Materials/](MaterialEditor/Materials/) | The NX material library files (`helion_material_library.xml`, `Gas_and_Liquid_helion_material_library.xml`, `OOTB/`, `backup/`). See [README Materials.txt](<MaterialEditor/Materials/README Materials.txt>) for how to point NX at this folder. |

## Requirements

- Windows with Excel (to edit the exported workbook)
- Python 3 (standard library only, no packages to install)

## Quick start

1. Run `MaterialEditor\export_to_excel.bat`.
2. Edit the workbook, save it and close Excel.
3. Run `MaterialEditor\import_to_xml.bat`. A timestamped backup of the XML is written to `MaterialEditor\Materials\backup\` first.
