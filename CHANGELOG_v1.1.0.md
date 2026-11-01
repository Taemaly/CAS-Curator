# CAS Curator v1.1.0 — Full Changelog

## Changes since v1.0.1

This changelog covers the implemented public v1.1.0 release only. Research-only and future-roadmap work is intentionally excluded.

---

# New

## Theme Colour Studio

A full **Theme Colour Studio** is now available in the public release.

It includes:
- live theme editing with a real-time preview
- preview click-to-jump to the related colour section
- custom theme creation, update, rename, delete, apply and management
- native `.s3theme` import/export
- **Apply This Theme** directly inside the Studio
- **Manage Themes…** inside the Studio, using the same management actions as the main Theme menu
- palette file import
- **Paste Link / HEX Colours…**
- drag-and-drop colour assignment
- **Swap**
- **Auto-fill**
- **Shuffle Unlocked**
- **Undo**
- **Revert Session**
- a persistent **My Colours** shelf
- **Readability Check** with jump-to-problem support
- Open Custom Themes Folder
- Reload Custom Themes
- Save as New Theme
- Update Theme

Built-in themes remain protected.

Release custom themes and My Colours use their own per-user Release storage. Compatible DEV custom themes/My Colours can be copied non-destructively on first use; after that, DEV and Release storage remain independent.

### Theme Colour Studio locks

Two lock systems are available:
- **Section Lock** — keeps that colour section fixed in place.
- **Shuffle Lock** — keeps a colour available during Shuffle Unlocked without permanently pinning it to one section.

Lock controls have stronger labels, larger indicators and theme-aware outlines for better visibility, including at 125% Windows scaling.

### My Colours improvements

My Colours now:
- saves automatically
- adds colours immediately with no separate Save step
- persists Remove and Clear actions immediately
- does not cause a theme save prompt when only the My Colours shelf changed
- provides a dedicated visual Remove Colour selector
- shows a clear selected state before removal
- confirms removal using the selected colour swatch and HEX value

Fixed an issue where removing a later saved colour could remove the first colour instead.

### Save-before-exit protection

Closing Theme Colour Studio after a real unsaved theme edit now asks whether to:
- Save
- Don't Save
- Cancel and return to the Studio

My Colours-only changes do not trigger this prompt because they already save automatically.

### Theme backup export

Added **Export Theme Backup…**.

It creates one dated ZIP containing:
- all custom themes under their saved names
- Theme Colour Studio metadata/locks
- My Colours when available
- a readable backup information file

The backup copies data only; it does not move or delete the live themes.

---

# Theme menu and theme-management improvements

The main **Theme** menu is now organised for faster access:
1. **Manage Themes**
2. **Theme Colour Studio…**
3. **Custom Themes**
4. built-in themes

**Manage Themes** contains palette/theme import, custom-theme management, backup, folder and reload actions.

Theme Colour Studio also has its own **Manage Themes…** button so those actions remain available without closing the Studio.

The Custom Theme Manager now shows a clear full-row outline around the selected theme so it is obvious which theme is about to be renamed, exported or deleted.

---

# Custom-theme colour behaviour

## Linked Library panel colours

Custom themes now use one linked colour for:
- **Toolbar panel**
- **Search / filter panel**

Those two Library surfaces always stay visually unified.

The actual **Toolbar / filter buttons** remain a separate editable/shuffleable colour.

The Theme Colour Studio preview uses the same mapping and allows clicking the Library header strip to jump to the linked panel colour role.

## Master Database path outline

Custom themes expose **Master Database path border** as its own Theme Colour Studio colour role.

Final behaviour:
- the path interior shares the surrounding Master Database panel colour
- only the thin outline has its own colour
- the outline can be edited manually
- it can be Auto-filled
- it can participate in **Shuffle Unlocked**
- it supports Section Lock and Shuffle Lock

**Aqua Depth** uses a green-blue/teal Master Database path outline.

## Automatic custom-theme text colours

Main and muted text are chosen automatically from readable light/dark variants based on the actual surface background.

Theme Colour Studio no longer exposes Main Text or Muted Text as editable colour sections.

Older custom themes that contain legacy text-colour keys remain compatible, but those legacy values no longer control the displayed text.

## Adaptive secondary/help text

Overview helper and experimental text now calculates a muted colour from the actual detail-card background, targeting readable contrast while remaining visually secondary.

The selected package details refresh immediately after switching/shuffling themes so the helper text cannot keep a stale colour from the previous theme.

---

# Identifier and ID lookup improvements

## Identifiers view

A dedicated **Identifiers** view separates technical identity information into clearer groups, including:
- package identity
- CAS/resource identity
- in-game lookup identity
- file/location evidence

Raw stored IDs remain visible as technical evidence without being presented as dependency-analysis results.

## Find ID

Added a compact **Find ID…** tool that can search technical identifiers across the Library, including values not currently shown as visible columns.

Lookup can use stored evidence such as:
- CASP/resource IDs
- package hashes
- technical resource identifiers
- stored paths/location evidence

Find ID:
- does not require a package to already be selected
- includes a Paste button
- clears temporary Quick Filters when necessary so a valid result is not hidden
- switches to Library
- selects the first result
- opens package details and Identifiers

Find ID is an **Advanced** feature and appears only on the **Library** page. It hides automatically on Scan and Guide.

## Identifier actions

Added:
- Copy Identifiers
- Copy CASP IDs
- a compact **Copy** button beside CASP ID(s) in CASP Observations
- Copy Exact Path
- Show Package in Explorer

The normal Library search is reserved for ordinary Library information such as package names, creators, types, Status, Profiles and notes. Technical IDs, hashes and stored paths are handled by **Find ID…**, keeping identifier lookup separate from live Library filtering.

---

# Performance Risk Score improvements

## Health Status and Performance Risk are clearly separated

User-facing wording now consistently distinguishes:
- **Health Status** — package/structural problems
- **Performance Risk Score** — advisory estimate of likely game-performance load/risk

A high Performance Risk Score does **not** mean that CC is broken or unsafe.

The internal MDB field name `Risk Score` is retained for compatibility, while the public UI uses **Performance Risk Score** or **Performance Score** where space is limited.

## Performance Risk calibration

### VPXY
Raw VPXY count no longer adds Performance Risk points.

VPXY remains available for technical information, structural Health checks and identifiers.

### 2K textures
A `2048 HIGH` texture warning now contributes **+1** instead of +2.

### 4K textures
4K scoring remains:
- one 4K texture: **+5**
- three or more 4K textures: **+1 additional point**

### Category-aware poly thresholds

Heavy / Very Heavy boundaries:
- Shoes: **7,000 / 12,000**
- Top: **7,000 / 12,000**
- Bottom: **7,000 / 12,000**
- Full Body: **13,000 / 20,000**
- Hair: **30,000 / 40,000**
- Accessories / small CAS parts: **7,000 / 12,000**
- Skin / Custom Sim / Morph Slider: no poly penalty
- Generic fallback: **15,000 / 30,000**

Heavy remains +3 and Very Heavy remains +5.

### GEOM size
GEOM scoring remains:
- below 1 MB: +0
- 1 MB to below 2 MB: +1
- 2 MB or more: +2

Final score remains capped at 10:
- 0–1: Safe
- 2–3: Moderate
- 4–5: High
- 6–10: Extreme

Health Status remains independent from this calculation.

---

# Refresh Performance Scores

Added/improved:

**Tools > Library / Maintenance > Refresh Performance Scores…**

It can refresh:
- selected Library packages
- the whole Library

This is intended for Libraries created with an older Performance Risk Score model.

## Preview before changes

The refresh workflow shows a preview before anything is written, including:
- Changed
- Already current
- Skipped
- source used for refreshed rows

Applying changes is explicit, and a normal Master Database backup is created before updates are written.

## Existing report and package fallback

For older Libraries, Refresh Performance Scores can use a compatible `Health_Scan_Report.xlsx` rather than requiring the original package files.

If a compatible report is unavailable, CAS Curator can locate a matching package using:
1. current File Location
2. configured CAS Curator managed File Locations
3. stored Original Source Path
4. an optional user-selected package root

Stored package hashes are verified before a physical package is used when a hash is available. Ambiguous matches are skipped rather than guessed.

## Automatic old-score detection

The Master Database now tracks an internal Performance Score model version.

When an older/missing model is detected on Library load, CAS Curator can offer to refresh the whole Library.

A partial/selected refresh does not incorrectly mark the entire Library as current.

## Clearer zero-change results

Refresh previews distinguish genuinely current packages from packages that could not be refreshed because no compatible source was available.

---

# Performance data stored in the Master Database

New scans now save the measurements required for future Performance Score recalculation directly in the Master Database:
- Texture Resolutions
- GEOM Total Bytes
- 4K Texture Count
- Performance Data Version

Older MDBs can be safely backfilled during a successful Performance Score refresh.

Once a row has the current Performance Data Version, future score recalculation can use the MDB directly instead of depending on the old Health Scan Report or physical package file for those measurements.

Backfilling can occur even if the visible score itself was already correct.

Health Status, classification, Last Scanned, Routing, notes and unrelated Library data are not changed by this migration.

---

# Health Scan Report is now optional

Added persistent option:

**Options > Create Health Scan Report after Scan**

It is ON by default.

When disabled:
- Scan still runs normally
- the console log still runs
- the Master Database still updates
- the separate Excel scan report is not created

Because the required Performance Score measurements are now stored in the MDB, the optional report is no longer required for future score refreshes on newly scanned data.

---

# In-app updater and changelog

Added:

**Tools > Help > Check for Updates…**

Stable public builds can:
- check the official GitHub Releases feed
- show current version → available version
- display release notes/changelog
- download and install a recognised Windows installer

Added:
- **Options > Check for updates on startup**

When already current, the startup check can remain quiet.

## Update safety

Automatic installation:
- accepts Windows `.exe` / `.msi` release assets only
- never executes ZIP/source assets
- requires an official matching SHA-256 checksum asset
- verifies installer size when available
- verifies SHA-256 before execution
- rejects non-GitHub/non-HTTPS update URLs
- deletes the partial installer if verification fails

Pending MDB changes and active Scan/Routing or Performance Score operations must be resolved before installation.

---

# Guide improvements

## Search Guide

Added:
- live Guide search
- highlighted matches
- Previous
- Next
- current-match counter

## Easier navigation and wording

The Guide now:
- puts **What CAS Curator Can Do** near the beginning
- includes clearer “Where to find it” directions for major features
- defines CASP before later technical use
- explains why CASP/resource IDs are useful
- defines Profile Membership earlier
- simplifies dense Table, Routing, Performance and Advanced explanations
- better explains ZIP/Sims3Pack intake
- clarifies CC Magic Import as a deployment/tracking copy location rather than the organised master location
- explains optional Health Scan Reports and MDB-stored performance data
- explains first Scan → Routing output/results folders more clearly
- includes Theme Colour Studio guidance matching the final UI

Inline technical examples such as `.package`, `.s3theme`, `__1` and `__2` now use ordinary clean text styling instead of grey code boxes.

## Theme-aware Guide colours

Theme Colour Studio includes:
- Guide Section
- Guide Highlight 1
- Guide Highlight 2
- Guide Highlight 3

Guide headings and callouts use those colours with automatic readability and stronger borders where needed.

---

# Library and UI improvements

## Whole-row hover and cleaner selection

Library hover now reads as a **whole-row state** across the frozen Package section and the scrollable columns.

Selected rows use:
- a subtle theme-aware tint
- thin top/bottom emphasis
- a strong left marker

This replaces the earlier heavy floating-box look.

Custom-theme row hover uses the actual row colour with automatic text contrast.

## Custom-theme Library row colouring

Added per-custom-theme option:

**Colour Library rows with this custom theme**

- ON: use the custom theme's Main table / Alternating table row colours
- OFF: use neutral white/grey Library rows while keeping the rest of the custom theme

The choice is stored in the `.s3theme` and preserved through export/import.

## Pinned packages

Pinned packages received a full usability pass:
- right-click **Pin This Item / Unpin This Item** action
- small pin icon in the package context menu
- small pin icon beside pinned Package names in the Library table
- compact one-line pin chips in the Pinned strip
- full filename shown on hover
- crowded chips preserve the **start** of the filename
- maximum of **5 pinned packages** per Master Database
- trying to pin a sixth shows a clear message
- existing saved pin lists are safely capped to the same five-item limit

The Pinned strip also shows a small pin icon beside its heading.

## Library right-click menu

The Library/Gallery package context menu now uses matching icons throughout for faster visual scanning.

The Pin/Unpin action uses the existing persistent pin system and stays in sync with the package-details Pin button.

## Switchable details panel

Added:

**Options > Details panel on left**

The Library package-details pane can now sit on either the left or right side.

CAS Curator remembers:
- chosen side
- splitter width
- collapsed/shown state

## Master Database row cleanup

The Opened Master Database identity area was refined so:
- the database caption/path stay grouped together
- Create/Open Master Database actions sit outside the identity strip
- custom-theme path interior shares the surrounding panel colour
- the thin path outline remains independently themeable

## About window

The About window now includes a small Ukrainian flag beside the creator information.

## Choose Visible Columns

Empty checkboxes now have a clearer theme-aware outline while retaining native checked-state rendering.

---

# Theme Colour Studio UI polish

Additional presentation/scaling improvements include:
- text + icon presentation for actions
- icons in Theme menus and Custom Theme Manager
- compact icon sizing
- sticky Auto-fill / Shuffle Unlocked / Undo / Revert Session controls
- sticky Save/Update/Close footer
- adaptive sticky-action width instead of a fixed width
- fixes for clipping at 125% Windows scaling
- palette heading moved to its own compact row
- Paste Link / HEX Colours available directly from the Studio
- colour-specific actions use small framed swatch icons
- widened colour-role editor for long role names
- Theme Colour Studio preview click-to-jump coverage expanded for the Library header strip

---

# Plain-language UI pass

High-impact dialogs were rewritten to put:
1. what happened
2. whether anything changed
3. what the user can do next

before implementation details.

Areas refined include:
- Performance Score refresh
- Master Database compatibility prompts
- Routing recovery
- managed Move/Copy
- ZIP/Sims3Pack safety messages
- Profile Layout save prompts
- save-conflict/failure messages
- File Location wording

Terminology cleanup includes:
- `MDB` written as **Master Database** in normal UI where clearer
- technical “write MDB” wording replaced by normal save wording
- unnecessary “preflight” wording removed from ordinary prompts
- updater UI explains successful verification without requiring users to interpret SHA-256 terminology

---

# Other UI/menu cleanup

- old Status migration/review action hidden from the normal public File menu
- outdated Guide wording about legacy/unverified Status removed
- fill-down option renamed **Confirm before fill-down prompt**
- standalone **Tools > Library / Maintenance > Preview / Extract ZIP** command removed
- ZIP handling remains available where it belongs: Scan and Routing Preview

---

# Fixed

Notable fixes since v1.0.1 include:
- My Colours Remove Colour could remove the wrong saved colour
- My Colours removal now has an explicit visual selector
- Theme Colour Studio startup/opening regressions found during testing
- Theme Colour Studio sticky actions could clip at 125% scaling
- palette header/instruction could clip at common Windows scaling
- custom-theme Library hover could become unreadable
- low-contrast selected rows could be hard to identify
- helper/experimental Overview text could be too dark or too light on shuffled/custom themes
- helper text could retain the previous theme's inline colour after switching themes
- whole-row hover did not read consistently across frozen and scrollable table sections
- Guide technical examples used distracting grey code-style boxes
- outdated/technical Guide and dialog wording
- visible literal `\n` formatting error in an Advanced Guide section
- zero-change Performance Score refresh messages could obscure missing-source/skipped cases
- automatic score-update prompt now explains why recalculation is recommended and that a full rescan is not needed
- selected-theme state in Custom Theme Manager could be unclear
- compact pinned chips could show the middle of filenames when crowded; they now preserve the beginning

---

# Compatibility and safety notes

- Existing MDB field key `Risk Score` is retained internally for compatibility.
- Existing older custom themes continue to load.
- Portable `.s3theme` files preserve lock data and unknown forward-compatible metadata.
- Release and DEV custom-theme/My Colours storage remain independent after first-use migration.
- Performance Score refresh does not rerun Health Status, classification, thumbnails or Routing.
- Performance updates do not change Last Scanned.
- Physical-package refresh fallback verifies stored hashes when available and skips ambiguous matches rather than guessing.
- Master Database backups are created before performance-data migration/score writes.
- External workbook changes are detected before overwrite.
- Health Status remains separate from Performance Risk Score throughout.

---

# Not included in v1.1.0

The following remain future work and are not advertised as v1.1.0 features:
- Resolve Used CC
- structured Saved Sim / SIMO dependency parsing
- Household/Lot dependency resolution
- Morph Dependency / Compatibility
- Package Cleanup / Reduce Size
- Merge Builder
- Store DBPP/EBPF support
- future expanded Random Sim Studio work
