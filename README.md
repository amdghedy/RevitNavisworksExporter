# Revit to Navisworks Batch Exporter

Batch-process Revit project files and create Navisworks cache files (.nwc) using Revit 2020 automation.

> **Compatibility note:** The current project is configured for Revit 2020, .NET Framework 4.8, and x64. Its source uses the standard installation path `C:\Program Files\Autodesk\Revit 2020\Revit.exe`. This is not yet a version-independent exporter.

## What it does

- Scans a selected folder for .rvt files.
- Exports each discovered model to a sibling .nwc file.
- Shows running status, progress, and a per-file log.
- Counts successful and failed exports and displays a final summary.
- Writes a timestamped text log in the selected source folder.
- Uses a Revit journal file to automate opening the model, invoking the Navisworks export command, and closing without saving.
- Offers a heuristic handling path for files that may be workshared.

## User interface and options

| Control | Behavior |
|---|---|
| **Browse** | Choose the folder containing Revit models. |
| **Include Subfolders** | Checked by default; search nested folders as well as the selected folder. When cleared, search only the selected folder. |
| **Show Revit Window** | Checked by default; run Revit visibly while processing. When cleared, request hidden-window processing. |
| **Start Export** | Process the discovered .rvt files one by one. |
| **Status / progress** | Displays the current file, processed count, and progress. |
| **Processing Log** | Shows file-level success/failure information. |

The exporter writes each NWC beside its source RVT, using the same base filename. The run log is named **NavisworksExport_Log_yyyyMMdd_HHmmss.txt**.

## Quick start: use the tool

1. Install Revit 2020 and confirm that the Navisworks exporter command is available in that Revit installation.
2. Build and deploy the project using the instructions below, or use a package produced by a trusted build of this repository.
3. Launch the batch exporter.
4. Select a folder containing source .rvt files.
5. Review the discovered file count.
6. Choose whether to include subfolders and whether Revit should remain visible.
7. Click **Start Export** and allow each model to finish.
8. Review the final successful/failed counts and inspect the generated log.
9. Open the resulting NWC files in your downstream review workflow and verify representative models.

**Before a large batch:** test on copies of a few models. Make sure output NWC files will not overwrite valuable existing deliverables.

## Build from source

### Requirements

- Windows
- Visual Studio 2022 with .NET desktop development tools, or compatible MSBuild tooling
- .NET Framework 4.8 targeting pack
- Revit 2020 installed, including RevitAPI.dll and RevitAPIUI.dll
- Navisworks exporter integration available in Revit 2020

### Build steps

1. Clone or download this repository.
2. Open BatchNavisworksExporter.csproj in Visual Studio.
3. Confirm that the Revit API references resolve. The current project references:
   - C:\Program Files\Autodesk\Revit 2020\RevitAPI.dll
   - C:\Program Files\Autodesk\Revit 2020\RevitAPIUI.dll
4. Build the project for **x64** using **Release** configuration.
5. Review the build output and the configured post-build deployment destination before running it.

The project contains a post-build copy step targeting %APPDATA%\Autodesk\Revit\Addins\2020. Because that step is path-specific, inspect it before building on a different workstation. Do not assume a successful compilation proves the export automation works in every environment.

## How the automation works

1. Enumerates .rvt files in the selected folder, optionally recursively.
2. Creates a temporary Revit journal for each model.
3. Launches Revit 2020 with the journal.
4. Attempts to open the model, invoke the Navisworks export command, save the NWC beside the RVT, close the model without saving, and exit Revit.
5. Waits up to ten minutes for each Revit process.
6. Checks whether an NWC exists and whether its size is greater than 1 KB.
7. Records the result and proceeds to the next file.

The workshared-file detection is only a **filename/size heuristic** (for example, certain filename patterns or files larger than 50 MB); it is not a reliable Revit worksharing metadata check.

## Troubleshooting

- **Revit not found:** install Revit 2020 at the configured default path, or update the source to use your actual installation path before building.
- **Export command not found / journal does not complete:** verify the Navisworks exporter integration in Revit 2020 and manually test an NWC export from Revit first.
- **No NWC generated:** inspect the per-run log and try the source model manually; journal UI automation can be sensitive to dialogs and Revit configuration.
- **Timeout:** the current per-file timeout is ten minutes. Large or slow models may exceed it.
- **NWC exists but is invalid:** the current validation only checks file existence and a minimum size, not full NWC integrity.
- **Access denied or files unavailable:** check write permissions, file locks, network paths, and whether another Revit process is using the model.
- **Workshared model prompts:** the current handling is heuristic and journal-driven. Validate detach/open behavior on a copy before processing central models.

## Known limitations and safety

- Targets Revit 2020 and a hard-coded default installation path.
- Relies on Revit journal/UI automation and the exporter command being present.
- The file-size check is a basic sanity check, not a full validation.
- The application should be tested against representative models and project standards before production use.
- Work on copies of important models and verify output files independently.

## Repository

- Project: BatchNavisworksExporter.csproj
- Main source and UI: Program.cs
- Revit add-in manifest: BatchNavisworksExporter.addin
- Solution and source: browse the repository files.
