# CAS Curator v1.1.0

CAS Curator is a specialist **The Sims 3 CAS catalogue, review and organisation tool**.

It works with a CC Master Database (MDB) and the CAS scanner to help inspect, classify, review and organise Sims 3 CAS custom content.

## Download

The recommended download is the **CAS Curator v1.1.0 Windows installer** from the v1.1.0 GitHub release.

Release packages:

- **`CAS_Curator_v1.1.0_Setup.exe`** — direct Windows installer; also used by CAS Curator's in-app updater.
- **`CAS_Curator_v1.1.0_Installer.zip`** — the same installer inside a ZIP for sites/downloads that prefer archived files.
- **`CAS_Curator_v1.1.0_APP.zip`** — portable compiled application.
- **`CAS_Curator_v1.1.0_SOURCE.zip`** — complete v1.1.0 application/source reference package.
- **`CAS_Curator_v1.1.0_THIRD_PARTY_SOURCE.zip`** — corresponding third-party source/licence bundle.

## What's new in v1.1.0

v1.1.0 is a major usability, theme-customisation and maintenance update.

Highlights include:

- **Theme Colour Studio** with live preview, click-to-jump sections, Auto-fill, **Shuffle Unlocked**, locks, Undo/Revert, My Colours, readability checks, import/export and **Apply This Theme**
- clearer **Manage Themes** tools and dated theme-backup ZIP export
- **Find ID…** for technical CASP/resource IDs, hashes and stored paths, including identifiers that are not visible as Library columns
- a dedicated **Identifiers** view with package, CAS/resource, in-game lookup and location/source evidence separated clearly
- a compact **Copy** button beside CASP ID(s) in CASP Observations
- clearer separation of **Health Status** from **Performance Risk Score**
- recalibrated category-aware Performance Risk scoring and **Refresh Performance Scores…** for existing Libraries
- full-row Library hover/selection improvements and stronger dark/custom-theme readability
- compact one-line pinned-package chips, visible pin icons and direct Pin/Unpin actions
- switchable left/right Library Details pane
- searchable built-in Guide, optional Health Scan Report creation and stable-release update checking
- numerous scaling, theme, wording and UI refinements

See [`CHANGELOG_v1.1.0.md`](CHANGELOG_v1.1.0.md) for the complete user-facing list of changes since v1.0.1.

## Search and technical identifiers

The normal Library **Search** is for ordinary Library information such as package names, creators, types, Status, Profiles and notes.

For technical identifiers — including CASP/resource IDs, package hashes and stored paths — enable **Advanced features** and use **Find ID…**. Find ID can search hidden technical evidence, clear temporary filters when needed, select the matching package and open its identifier information.

## Performance Risk Score

**Health Status** and **Performance Risk Score** are intentionally separate:

- **Health Status** reports supported package/structural problems.
- **Performance Risk Score** is an advisory estimate of likely game-performance load/risk based on supported evidence such as polycount, GEOM size, texture resolution/quantity and category-aware thresholds.

A high Performance Risk Score does not mean that CC is broken or unsafe.

## Safety

- Run Scan does not move source files.
- Run Routing is an explicit action and can be previewed before files are moved.
- Scanner writes create MDB backups.
- Saving a Planned Location does not move files.
- Applying Profile Locations requires its own explicit previewed action.
- CC Magic Import uses copies and does not replace the organised master package.
- ZIP extraction keeps the original archive, rejects unsafe paths and does not silently overwrite existing files.

Keep normal backups of important MDBs and custom content.

## Running from source

For the complete source package:

1. Extract the whole CAS Curator v1.1.0 source folder to a normal writable folder.
2. Run `run_CAS_Curator_v1.1.0.bat`.
3. CAS Curator uses a private `.venv` in its own folder.
4. A supported 64-bit Python 3.10–3.14 installation is required.
5. The published v1.1.0 source launcher uses the tested **PySide6 / Qt 6.12.0** runtime.

The compiled Windows build additionally uses a separated Python 3.8 scanner/helper architecture. The complete tested Windows build kit is kept separately from the public source package.

## Source and licence

The source is published for visibility and reference, but **CAS Curator is proprietary software, not open-source software**.

Copyright © 2026 Taemaly. All Rights Reserved.

Copying, modification, redistribution, reuse, mirroring, or creation of derivative works is not permitted without prior written permission.

Third-party components remain under their own licences. See the included notices and the corresponding third-party source bundle.

See [`LICENSE`](LICENSE) for the full CAS Curator licence terms.
