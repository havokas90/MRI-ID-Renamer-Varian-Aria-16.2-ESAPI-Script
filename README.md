MRI ID Renamer — Varian Aria 16.2 ESAPI Script
A write-enabled ESAPI plugin script for Varian Aria 16.2 that automates renaming of imported MRI image IDs in Eclipse.
When MRI studies are imported for fusion planning, images arrive with placeholder IDs like 1 or MRI_1. This script reads the existing series description and study description already stored in Aria and generates a meaningful, standardised image ID — without any manual typing.
What it does:

Scans all MRI series for the currently open patient
Reads the series description (e.g. t2_tse_ax_p3_2.5mm) and strips slice thickness to save characters
Reads the study description (e.g. Pelvis Prostate) to automatically append a body region suffix (_PEL, _HN, _CH, _AB, _BR, _SP, _EXT)
Proposes new IDs within Aria's 16-character limit, truncating at word boundaries
Shows a preview window with checkboxes — select only the images you want renamed, useful when re-importing into an existing course
Applies changes with a single click after confirmation

Example:
Before:  1
After:   T2_TSE_AX_P3_PEL
Requirements:

Eclipse / Aria 16.2
Eclipse Scripting API license
Eclipse Automation license (for write access on clinical systems)
Visual Studio (to compile as a binary plugin)

The script is heavily commented throughout and written to be readable by anyone learning ESAPI — each section explains what it does, why, and how to extend it.

Licence & Disclaimer
© 2025 David Neill. All rights reserved.
This script is provided for educational and non-commercial use only. It was developed for internal use at [Your Hospital/Health Service] and is shared publicly to support learning within the radiation therapy and medical physics community.
You are welcome to:

Read, learn from, and adapt this code for your own non-commercial clinical or research use
Share it with attribution

You are not permitted to:

Use this script or any derivative of it for commercial purposes
Sell, license, or redistribute this script as part of a commercial product or service
Remove or alter this notice

Clinical Disclaimer:
This script is provided as-is, without warranty of any kind. It has been tested in a specific clinical environment and may require modification for use at other sites. The authors of custom scripts are responsible for verifying the accuracy and correctness of their scripts (Varian SYRS350.75). Always test on a non-clinical system before deploying to a clinical environment. The author accepts no liability for any clinical, technical, or data consequences arising from the use of this script.
This project is not affiliated with, endorsed by, or supported by Varian Medical Systems / Siemens Healthineers.
