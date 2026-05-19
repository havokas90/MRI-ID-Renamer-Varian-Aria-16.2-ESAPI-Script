# ImageIdRenamer - Varian ARIA ESAPI Plugin Script

Write-enabled ESAPI plugin script for Varian ARIA/Eclipse that renames imported MR and CT image IDs from the currently open patient.

The script builds proposed image IDs from the existing ARIA study and series descriptions, shows a preview table, and only writes changes after the user confirms the selected rows.

## What It Does

- Scans the open patient for MR and CT series.
- Uses `study.Comment` and `series.Comment` to build meaningful image IDs.
- Skips 2D scout/localiser images by requiring `image.ZSize > 1`.
- Removes slice-thickness fragments such as `_2.5mm`.
- Handles `ep2d`/DWI descriptions by keeping the meaningful final token where possible.
- Adds body-site suffixes such as `_PEL`, `_HN`, `_CH`, `_AB`, `_BR`, `_SP`, and `_EXT`.
- Adds a `CT_` prefix for CT images unless the scanner description already starts with CT.
- Keeps final IDs within ARIA's 16-character limit.
- Forces final IDs to uppercase and changes spaces to underscores.
- Detects existing patient image IDs and auto-numbers duplicate proposals.
- Shows a preview window with MR rows tinted blue and CT rows tinted orange.
- Lets users untick rows they do not want renamed.
- Lets users manually edit each proposed **New ID** before clicking Apply.
- Validates edited IDs for blank values and duplicate selected New IDs before writing.

## Which File To Use

Use this file in Varian Script Wizard / Script Builder:

`InstallReady/ImageIdRenamer.cs`

The root `ImageIdRenamer.cs` is the editable source copy. `InstallReady/ImageIdRenamer.cs` is kept in sync as the deployment copy.

## Script Wizard / Script Builder

Create the script as:

- Type: **Plugin**
- Language: **C#**
- Write access: **enabled / writeable**
- Platform: **x64**
- Framework: **.NET Framework 4.8** or the framework used by your local ARIA/ESAPI install

Do not create this as an **Executable**. This script is intended to run inside Eclipse from the currently open patient.

Keep the assembly attributes in the source file unless your Script Builder project generates a separate `AssemblyInfo.cs`. Do not duplicate the same assembly attributes in both places.

The write-enabled marker must remain:

```csharp
[assembly: ESAPIScript(IsWriteable = true)]
```

## Local Preview

This repo includes a small local WPF preview harness so you can inspect the preview window before deploying at work.

Run:

`Launch-LocalPreview.bat`

Local preview mode uses sample MR/CT rows. It does not open an ESAPI patient and cannot write image IDs.

## Varian Script Wizard Launcher

On a machine with Varian RTM 17.0 installed at the standard path, run:

`Launch-ScriptWizard.bat`

This opens:

`InstallReady/ImageIdRenamer.cs`

with:

`C:\Program Files\Varian\RTM\17.0\esapi\ScriptWizard.exe`

If your Varian install path differs, edit `Launch-ScriptWizard.ps1`.

## Validation Build

The `Validation` project compiles the single-file ESAPI script against local Varian ESAPI DLLs:

- `VMS.TPS.Common.Model.API.dll`
- `VMS.TPS.Common.Model.Types.dll`

Expected default path:

`C:\Program Files\Varian\RTM\17.0\esapi\API`

Run:

```powershell
dotnet build .\Validation\ImageIdRenamer.Validation.csproj -v:minimal
```

The validation build checks C# compile compatibility against the real local Varian API assemblies. It does not run against a patient.

## Clinical Use Notes

This script writes to ARIA by setting `image.Id`, so it requires:

- Eclipse / ARIA with ESAPI installed
- Eclipse Scripting API licence
- Eclipse Automation licence for write access on clinical systems
- Script approval for write-enabled use in your environment
- Local testing on a non-clinical patient or research/test system before clinical deployment

Only selected rows are written. If a user edits a proposed New ID, the script normalizes it to uppercase, replaces spaces with underscores, strips unsupported characters, enforces the 16-character limit, and blocks Apply if selected rows contain duplicate New IDs.

## Project Boundary

This repository is only for ImageIdRenamer. It should not contain PlanChecker code, PlanChecker launchers, PlanChecker packages, or PlanChecker documentation.

## Licence & Disclaimer

Copyright (c) 2026 David Neill.

Provided for educational and non-commercial use. You are responsible for local validation, commissioning, and clinical governance before any clinical deployment.

This script is provided as-is, without warranty of any kind. It was developed for a specific clinical environment and may require modification for other sites. The author accepts no liability for clinical, technical, or data consequences arising from use of this script.

This project is not affiliated with, endorsed by, or supported by Varian Medical Systems or Siemens Healthineers.
