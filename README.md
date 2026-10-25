# CAS Curator v1.0.1

CAS Curator is a specialist **The Sims 3 CAS catalogue, review and organisation tool**.

It works with a CC Master Database (MDB) and the validated CAS scanner to help inspect, classify, review and organise Sims 3 CAS custom content.

## Download

The recommended download is the **CAS Curator v1.0.1 installer** from the [v1.0.1 release](https://github.com/Taemaly/CAS-Curator/releases/tag/v1.0.1).

Release packages:

- **`CAS_Curator_v1.0.1_Installer.zip`** — recommended Windows installer. Lets you choose the install location, creates shortcuts, includes an uninstaller, and can launch CAS Curator when setup finishes.
- **`CAS_Curator_v1.0.1_APP.zip`** — portable compiled application for users who prefer to extract and run the app without installing it.
- **`CAS_Curator_v1.0.1_SOURCE.zip`** — complete v1.0.1 source package for reference and source-based use.
- **`CAS_Curator_v1.0.1_THIRD_PARTY_SOURCE.zip`** — corresponding third-party source/licence bundle.

## What's new in v1.0.1

v1.0.1 keeps the **frozen scanner core unchanged** while substantially expanding the organisation and review workflow around it.

Highlights include:

- Profile Rules that remember each Profile's filter view and visible Table columns
- safer Profile Membership and Experimental Profile Layout workflows with explicit Planned Locations
- read-only and conflict preflight for Apply Profile Locations, plus stronger rollback cleanup
- the Experimental Routing Structure Builder with reusable folder structures, preview, sanitising and transactional routing
- managed File Locations and safer routing/recovery location updates
- universal **+ Add Quick Filter…** support alongside Status / CC Type / Creator
- Creator and Converter metadata plus improved package web-search shortcuts
- persistent column visibility, Select all / Deselect all / Default controls, collapsible column shortcuts and smoother horizontal Table navigation
- a hideable Details panel and compact Gallery quick-filter toolbar
- Age/Gender plus opt-in Advanced and Experimental metadata views
- resizable Scanner Log, clearer Tools organisation, improved dialogs and updated Guide
- **About CAS Curator…**, Copy version info, source/support shortcuts and final application icon polish

Additional convenience tools include copying selected packages to a configured CC Magic Import folder, ZIP preview/extraction, and conservative Sims3Pack scan intake.

See [`CHANGELOG_v1.0.1.txt`](CHANGELOG_v1.0.1.txt) for the full user-facing list of changes since v1.0.

## Safety

- The frozen scanner core used by this release is unchanged.
- Run Scan does not move source files.
- Run Routing is an explicit action and can move source files.
- Routing duplicate handling was corrected so a package is not treated as a duplicate of its own freshly scanned MDB record.
- Risk Score is a review value and does not decide routing folders.
- Experimental scanner-derived metadata is for review only and never automatically controls Routing.
- Experimental Profile Layouts are separate from scanner metadata. Saving a Planned Location stores only a plan; package files move only through the explicit previewed Apply Profile Locations action.
- Copying selected packages to the configured CC Magic Import folder uses copies and does not replace the organised master package.
- ZIP extraction is explicit, keeps the original ZIP, rejects unsafe paths and does not overwrite existing files.
- RAR and 7Z preview/extraction are intentionally not supported.

## Running from source

For the complete source package:

1. Extract the whole CAS Curator v1.0.1 source folder to a normal writable folder.
2. Run `run_CAS_Curator.bat`.
3. CAS Curator uses a private `.venv` in its own folder.
4. A supported 64-bit Python 3.10–3.14 installation is required.
5. On first start, PySide6 may be installed into that private environment.

Your existing Sims scanner Python installation can remain unchanged.

## Important

Keep backups of important MDBs and custom content. CAS Curator includes safety checks, previews and MDB backups, but any tool that can move files should still be used with normal backup precautions.

## Source and license

The source is published for visibility and reference, but **CAS Curator is proprietary software, not open-source software**.

Copyright © 2026 Taemaly. All Rights Reserved.

Copying, modification, redistribution, reuse, mirroring, or creation of derivative works is not permitted without prior written permission.

See the [`LICENSE`](LICENSE) file for the full terms.
