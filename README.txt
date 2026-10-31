CAS CURATOR v1.1.0 — SOURCE PACKAGE
====================================

CAS Curator is a specialist The Sims 3 CAS catalogue, review and organisation tool.
This archive contains the proprietary CAS Curator v1.1.0 application source plus the
runtime/reference files needed to inspect the published source package.

IMPORTANT
---------
CAS Curator is proprietary software and is not open-source software.
Copyright © 2026 Taemaly. All Rights Reserved.
See LICENSE.txt for the complete licence terms.

The S3PI/s3pe components included under helper/ are third-party GPL components and
are covered by their own licence and acknowledgement files. Their corresponding
upstream source is distributed separately in:
CAS_Curator_v1.1.0_THIRD_PARTY_SOURCE.zip

WHAT'S NEW IN v1.1.0
--------------------
Major additions and refinements include:
- Theme Colour Studio with live preview, locks, Shuffle Unlocked, My Colours,
  readability checks, import/export and Apply This Theme;
- Find ID and clearer Identifiers views;
- clearer separation of Health Status and Performance Risk Score;
- recalibrated Performance Risk scoring and previewed score refresh;
- improved Library hover/selection and custom-theme readability;
- compact pinned-package chips with tooltips and direct Pin/Unpin actions;
- switchable left/right Library Details pane;
- improved Theme management, Guide search/navigation and updater support;
- optional Health Scan Report creation;
- many scaling, wording, theme and UI fixes.

See RELEASE_NOTES.md and CHANGELOG.md for the complete public release information.

SOURCE-PACKAGE START
--------------------
1. Extract the complete source package to a normal writable folder.
2. Run run_CAS_Curator_v1.1.0.bat.
3. A supported 64-bit Python 3.10–3.14 installation is required for the GUI source build.
4. The launcher creates a private .venv in the CAS Curator folder and installs the
   required PySide6 package there if needed.

The public Windows installer/compiled build is the recommended download for normal users.

SCANNER / THIRD-PARTY NOTE
--------------------------
The scanner reference source uses the historical S3PI/s3pe libraries. The application
source package carries the same EA reference data and third-party helper material used
by the v1.0.1 source publication, while v1.1.0 updates the scanner-side Performance
Risk measurements/model data stored in new scans.

KEEP BACKUPS
------------
CAS Curator includes previews, safety checks and Master Database backups, but users
should still keep normal backups of important custom content and Master Databases.
