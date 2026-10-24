CAS CURATOR v1.0.1
==================

CAS Curator is a specialist Sims 3 CAS catalogue, review and organisation tool.
It works with a CC Master Database (MDB) and the validated CAS scanner to help
inspect, classify, review and organise Sims 3 CAS custom content.

WHAT'S NEW IN v1.0.1
--------------------
This update keeps the frozen scanner core unchanged while substantially improving
profiles, Library filtering, managed locations, routing safety and day-to-day UI.

Highlights include:
- Profile Rules that remember each Profile's filter view and visible Table columns;
- safer Profile Membership and Profile Layout workflows with explicit Planned Locations;
- read-only and conflict preflight for Apply Profile Locations, plus stronger rollback cleanup;
- the Experimental Routing Structure Builder with reusable folder structures, preview, sanitising and transactional routing;
- managed File Locations and safer routing/recovery location updates;
- universal + Add Quick Filter… support alongside Status / CC Type / Creator;
- Creator and Converter metadata plus improved package web-search shortcuts;
- persistent column visibility, Select all / Deselect all / Default controls, collapsible column shortcuts and smoother horizontal Table navigation;
- a hideable Details panel and a compact Gallery quick-filter toolbar;
- CC Magic Import deployment-copy workflow with live status;
- ZIP preview/extract and conservative Sims3Pack scan intake;
- Age/Gender plus opt-in Advanced and Experimental metadata views;
- resizable Scanner Log, clearer Tools organisation, improved dialogs and updated Guide;
- About CAS Curator…, Copy version info and source/support shortcuts.

See CHANGELOG_v1.0.1.txt for the full user-facing list of changes since v1.0.

SAFETY
------
- The frozen scanner core used by this release is unchanged.
- Routing duplicate handling was corrected so a package is not treated as a duplicate of its own freshly scanned MDB record.
- Risk Score is the public review value and does not decide routing folders. Legacy textual Risk data may remain preserved in older MDBs for compatibility but is not exposed as a public Quick Filter.
- Automatic scanner thumbnails live in a Thumbnails folder beside the MDB; when an MDB and that folder are moved together, CAS Curator can resolve matching hash-named thumbnails from the new location. Manual preview-image links are left unchanged.
- Experimental scanner-derived metadata is for review only and never automatically controls Routing.
- Experimental Profile Layouts are separate from scanner metadata. Saving a Planned Location stores only a plan; package files move only through the explicit previewed Apply Profile Locations action.
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

Profiles and filter views
- Current Profile is the working Profile selector. Changing it never moves package files.
- Manage Profiles… handles create/rename/delete, Profile Membership context, Profile Rules and Experimental Profile Layouts.
- The permanent Library filters are Status / CC Type / Creator. + Add Quick Filter… adds temporary extra filters for other columns.
- Profile Rules save the current Library filter view and visible Table columns for that Profile.
- Switching Current Profile clears the previous profile-view filters, then applies the destination Profile's saved rules if it has any.
- Clear Rules removes the saved Profile Rules and clears the rule-driven filter view; Search text is separate and is not saved as part of Profile Rules.
- Profile Membership is global package metadata. A package can belong to multiple Profiles and switching Current Profile never erases or replaces those memberships.
- Profile Rules never move files or change Profile Membership.

Profile Location table editing:
- Current File Location is read-only in the Library Table; it represents physical truth now.
- With Experimental features enabled, Planned Location is a profile-specific Table column.
- A Planned Location cell is editable only when a Current Profile is selected and that package belongs to it, matching the Edit-tab rules.
- Saving a planned location never moves files; Apply Profile Locations is still required for physical movement.

Profile-specific visible columns:
- Manage Profiles > Profile Rules > Save Current View saves the Current Profile's Quick Filters and visible Table columns.
- Different Profiles can remember different visible-column sets.
- Switching to a Profile restores its saved columns when Profile Rules exist; older rules without saved columns keep the current column layout.
- Search text is still not saved, and Profile Rules never move files or change Profile membership.

Table drag-fill:
- Drag-fill remains available on editable Table cells.
- Dragging from a populated source copies that value down across the displayed rows.
- Current File Location is intentionally read-only in Table view, so drag-fill does not apply to it.
- Planned Location supports drag-fill only for rows that meet the same Current Profile/membership conditions as the Edit tab.

Table bulk clear:
- Populated drag-fill remains available for copying values down.
- To clear one editable field across several packages, Ctrl/Shift-select rows in that column and press Delete.
- Clearing always asks for confirmation and remains staged until Save Changes.
- Shift+Delete stages Library-record deletion; physical package files are never deleted by that action.
- Current File Location remains read-only; Planned Location clears only on eligible Profile rows.

Profile membership safety:
- The Table column is labelled Profile Membership because it shows all Profiles the package belongs to; it does not follow Current Profile.
- Double-clicking Profile Membership and choosing a Profile now adds that membership instead of replacing existing memberships.
- Current Profile controls Profile-specific Planned Location / Profile Rules context.

Removing Profile membership:
- The Edit tab now has a Remove from dropdown containing only Profiles assigned to the focused package.
- Removal is confirmed and staged until Save Changes.
- If that Profile has a Planned Location for the package, the warning explains that saving the membership removal also removes that Profile-specific plan.

Profile Add/Remove dropdowns:
- Add to shows only Profiles the focused package is not already assigned to.
- Remove from shows only Profiles the focused package is already assigned to.
- Both lists refresh immediately when membership changes are staged.

Profile Membership meaning:
- Membership means the package belongs to that Profile and can have that Profile's own Planned Location.
- A package can belong to multiple Profiles, each with a different Planned Location.
- Membership alone never moves a file; Apply Profile Locations uses the Profile's membership plus its saved Planned Locations to carry out planned moves.

Experimental Routing Structure Builder:
- Choose a managed destination root, then build a literal folder hierarchy from MDB fields (for example Creator > CC Type > Subtype).
- Folder levels can be dragged/reordered; custom fixed folder labels are supported.
- Missing values can use Unassigned, skip the level, or block the package.
- Preview shows exact destinations before any move.
- Route with Structure preflights the whole batch, never overwrites, copy/verifies SHA-256, then deletes the original with rollback on failure.
- Structures can be saved as reusable presets.
- The routed result can optionally be saved as a Profile (membership + destination-root Planned Location).
- Setting a non-blank Planned Location manually now automatically adds the Current Profile to Profile Membership if needed.
