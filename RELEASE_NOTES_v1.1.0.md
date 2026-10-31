# CAS Curator v1.1.0 — Release Notes

v1.1.0 is a major usability and theme-customisation update, with additional work on identifiers, Performance Risk scoring, Library readability and maintenance tools.

## Highlights

### Theme Colour Studio
Create and edit custom themes with a live preview, click-to-jump colour sections, Auto-fill, Shuffle Unlocked, locks, Undo/Revert, My Colours, readability checks, import/export and **Apply This Theme** directly from the Studio.

Custom-theme management is now easier to reach from both the main Theme menu and inside the Studio. You can also export a dated backup ZIP containing all custom themes and My Colours.

### Find ID and Identifiers
The new **Find ID…** tool searches technical identifiers, hashes and stored paths even when those fields are not visible as Library columns. The Identifiers view separates package identity, CAS/resource identity, in-game lookup identity and location/source evidence.

### Performance Risk improvements
Health Status and Performance Risk Score are now clearly separated. The Performance Risk model received category-aware poly recalibration, lighter 2K texture weighting and removal of raw VPXY-count scoring.

Existing Libraries can use **Refresh Performance Scores…** with a preview before changes. New scans store the measurements needed for future Performance Score recalculation directly in the Master Database.

### Library and theme readability
Library hover now highlights the whole row. Selected rows use a cleaner theme-aware treatment. Custom-theme helper text adapts automatically to the actual background colour, including after shuffling or switching themes.

Custom themes now link the Toolbar panel and Search/filter panel to one shared surface colour while keeping Toolbar/filter buttons separately shuffleable.

### Pinned packages
Pinned items now use compact one-line chips with full-name tooltips, a five-pin limit and visible pin icons in both the Pinned strip and Library Package column. The right-click package menu includes direct Pin/Unpin actions and matching icons throughout.

### Other improvements
- switch the Library details pane between left and right
- reorganised Theme menu and Manage Themes workflow
- optional Health Scan Report creation
- in-app update checking for stable releases
- searchable built-in Guide with improved wording/navigation
- Theme Colour Studio and Guide scaling/polish improvements
- small Ukrainian flag added to About

For the complete list, see `CHANGELOG.md`.

## First launch

Run:

`run_CAS_Curator_v1.1.0.bat`

CAS Curator uses a private Python environment in its own folder. The first launch may take longer while required GUI components are installed; later launches are normally much faster.

### Windows console note

During a first-time install, avoid clicking/dragging inside the black Command Prompt window. On some Windows setups this can enter Select mode and pause the installation. If that happens, press **Esc** to leave Select mode and allow startup to continue.
