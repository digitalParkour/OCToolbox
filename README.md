# OC Toolbox v.1.1

---

## PURPOSE / ISSUE TO SOLVE

| Tool        | Description |
| ----------- | ----------- |
| Report Tool | Pick properties and dump to csv. Get all beyond page size. Gather parent and child data together. |
| Export Tool | Guided OC seed export for prompted marketplace. |
| Import Tool | Extend OC seed mechanism to target any environment. |
| Delete Tool | Quickly purge data for selected types. |

---

## USAGE

Download the OCToolbox.ps1 file.

Using a powershell terminal (preferably windows terminal) run the script file:

```PowerShell
 .\OCToolbox.ps1
```

The rest is guided by the script.

Demo: [Video Tour](https://clipchamp.com/watch/7CTIVyxEI0L) ~16 min

---

## PRE-REQS

The import tool requires the **PSYaml** module to parse yml files.

 - Goto: https://github.com/Phil-Factor/PSYaml
 - Copy files to %USERPROFILE%\Documents\WindowsPowerShell\Modules\PSYaml

If you download the files to a different folder, edit this line near the top of OCToolbox.ps1:

```PowerShell
 $PSYamlFolder = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\PSYaml"
```

---

## RELEASES

### v1.1 - 2023.08.24

Improvements:

 - Report Tool - Added support for reporting on nested types. Can now list parent columns along side child entities. ie. Product.ID, Variant.ID, etc.
 - ChoiceMenu - now handles case of list being longer than window gracefully

### v1.0 - 2023.08.23

Initial release

Features:

 - Report Tool
 - Export Tool
 - Import Tool
 - Delete Tool

---

## ROADMAP / places to go next

 - Report Tool - Support filters
 - Delete Tool - Support sub types (leverage nested Reporting work)
 - Delete Tool - Support deleting from list of ids
 - Delete Tool - Support deleting from browse and select
 - Import Tool - Support partial import from choice menu of types (honoring dependencies and api client id mapping) - current workaround is to edit seed file

 ---
