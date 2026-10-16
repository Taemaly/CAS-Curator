# CAS Curator v1.0.1

CAS Curator is a specialist **The Sims 3 CAS catalogue, review and organisation tool**.

It works with a CC Master Database (MDB) and the validated CAS scanner to help inspect, classify, review and organise Sims 3 CAS custom content.

## What's new in v1.0.1

This update keeps the validated scanner/routing core unchanged while improving:

- Library organisation and filtering
- Creator metadata and search tools
- File Location management and safer routing
- CC Magic Import deployment copies
- ZIP and Sims3Pack intake
- Age/Gender display from CASP observations
- Advanced and Experimental metadata views
- Table column visibility and editing tools
- Library layout and panel resizing
- Themed dialogs and long-path wrapping
- Dark-theme readability
- Expanded Scanner Log readability
- Guide and Manual fields help

See `CHANGELOG_v1.0.1.txt` for the full list of changes since v1.0.

## Safety

- Run Scan does not move source files.
- Run Routing is an explicit action and can move source files.
- Experimental metadata is for review only and does not automatically control Routing.
- CC Magic Import uses deployment copies and does not replace the organised master package.
- ZIP extraction is explicit, keeps the original ZIP, rejects unsafe paths, and does not overwrite existing files.
- RAR and 7Z preview/extraction are intentionally not supported.

## Running from source

1. Extract the whole CAS Curator folder to a normal writable folder.
2. Run `run_CAS_Curator.bat`.
3. CAS Curator uses a private `.venv` in its own folder.
4. A supported 64-bit Python 3.10–3.14 installation is required.
5. On first start, PySide6 may be installed into that private environment.

Your existing Sims scanner Python installation can remain unchanged.

## Important

Keep backups of important MDBs and custom content. CAS Curator includes safety checks, previews and MDB backups, but any tool that can move files should still be used with normal backup precautions.

## License

Copyright © 2026 Taemaly. All Rights Reserved.

CAS Curator and its source code are proprietary. Copying, modification, redistribution, reuse, or creation of derivative works is not permitted without prior written permission.

See the `LICENSE` file for full terms.
