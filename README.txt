CAS CURATOR v1.0.1
==================

CAS Curator is a specialist Sims 3 CAS catalogue, review and organisation tool.
It works with a CC Master Database (MDB) and the validated CAS scanner to help
inspect, classify, review and organise Sims 3 CAS custom content.

WHAT'S NEW IN v1.0.1
--------------------
This update keeps the validated scanner/routing core unchanged while improving
Library organisation, review tools, archive/Sims3Pack intake, CC Magic Import,
managed locations, filters, readability and the Guide.

Highlights include:
- universal + Add Filter support and a permanent Creator filter;
- manual Creator metadata and Creator-aware web-search shortcuts;
- File Location management, managed moves and safer routing/recovery updates;
- CC Magic Import deployment-copy workflow with live status;
- ZIP preview/extract and conservative Sims3Pack scan intake;
- Age/Gender display from direct CASP observations;
- opt-in Advanced and Experimental metadata views;
- persistent Table column visibility and safer editing tools;
- improved Library header grouping and restored-window panel resizing;
- themed compact dialogs with reliable long-text/path wrapping;
- improved dark-theme readability and a neutral grey expanded Scanner Log;
- updated Guide and simplified Manual fields help.

See CHANGELOG_v1.0.1.txt for the full user-facing list of changes since v1.0.

SAFETY
------
- The validated scanner/routing core used by this release is unchanged.
- Experimental metadata is for review only and never automatically controls
  Routing.
- Run Scan does not move source files.
- Run Routing is an explicit action and can move source files.
- CC Magic Import uses deployment copies and does not replace the organised
  master package.
- ZIP extraction is explicit, keeps the original ZIP, rejects unsafe paths and
  will not overwrite existing files.
- RAR and 7Z preview/extraction are intentionally not supported.

FIRST START / SOURCE PACKAGE
----------------------------
1. Extract the whole CAS_Curator_v1.0.1 folder to a normal writable folder.
2. Run run_CAS_Curator.bat.
3. CAS Curator uses a private .venv in its own folder.
4. A supported 64-bit Python 3.10-3.14 installation is required.
5. On first start, PySide6 may be installed into that private environment.

Your existing Sims scanner Python installation can remain unchanged.

IMPORTANT
---------
Keep backups of important MDBs and custom content. CAS Curator includes safety
checks, previews and MDB backups, but any tool that can move files should still
be used with normal backup precautions.
