# Material library Excel editor

Edits `MaterialEditor/Materials/helion_material_library.xml` through a spreadsheet: one row per material, one column per sub-property (`Name`, `Class`, `Subclass`, `MassDensityConstant_110`). Python 3 standard library only.

## Workflow
1. Run `export_to_excel.bat` - writes and opens `helion_material_library.xlsx`.
2. Edit the cells (not `#`, the label row and header row; don't insert/delete/reorder rows).
3. Save and close Excel.
4. Run `import_to_xml.bat` - validates all cells, backs up the XML to `MaterialEditor/Materials/backup/` with a timestamp, then writes only the changed values. If any cell is invalid nothing is written and the cells are listed.

`MassDensityConstant_110` is in kg/m^3 and is written back as `1.799194e+03` style. Only existing materials can be edited; adding/deleting materials is not supported.

## Other columns
`python material_xlsx.py export --props MassDensityConstant_110,Version_1` or `--all` (every property; the dependent tables appear as `<property> [values]` columns).
