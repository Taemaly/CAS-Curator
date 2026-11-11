# CAS Curator (CCS3) — Development Roadmap

> **Internal planning reference, not a public release promise.**
> **Updated:** 2026-11-11
>
> CAS Curator currently focuses on Sims 3 CAS content. Future improvements must not destabilise the existing CAS scanner, MDB, routing or public-release workflows. This file documents possible next work; it does not claim that every proposed CC type is currently supported.

## Planned feature: Detailed CC Type & Subtype Classification — ALL supported CC families

**Status:** Planned / deferred until the current CCS3 DEV scanner work is stable.  
**Priority:** Significant library-management enhancement, *not* part of the present scanner-fix batch.  
**Scope:** Expand accurate type/subtype recognition across **all categories of Sims 3 custom content**, not just Buy Mode, while retaining existing correctly identified CAS subtypes. For content families the scanner does not yet support, keep them explicitly in the roadmap for later parser support.

### Why

- Filter, find and organise specific items (e.g. sofas vs all seating, kitchen counters vs all surfaces, eyeliner vs all makeup).
- Improve optional subcategory-based routing and profile management without changing existing paths automatically.
- Separate CC identity from appearance, functionality, creator-style tags and technical health.
- Make database/library browsing and future randomiser/recommendation features more precise.

### User-facing classification structure

1. **CC Family / Mode:** CAS, Buy, Build, Create-a-Style / Patterns, Gameplay / Script / Tuning, and other supported Sims 3 content.
2. **Primary Type / Category:** Hair, Clothing, Accessory, Seating, Surface, Door, Window, etc.
3. **Subtype (optional):** Eyeliner, Lipstick, Sofa, Dining Chair, Archway, Bay Window, etc.
4. **Independent attributes:** age/gender/category flags, outfit categories, room tags, placement, style, functionality, object size, use, recolour/default replacement, expansion requirements, and other supported properties.

**Global fallback rule:** When a parent category is known but the subtype is not reliably identifiable, show only the **parent category** (e.g. `Fireplace`, `Chair`, `Hair`)—**no “Generic” prefix or redundant Generic subtype**. Use `Unknown` only where the broader category itself is uncertain. Never guess a precise subtype just to fill a field.

### Coverage — ALL CC types

The following are **coverage groups for investigation**, not guaranteed existing scanner outputs or a complete fixed taxonomy. Add missing groups when real Sims 3 CC examples require them.

| CC family | Detailed types/subtypes to support or investigate |
| --- | --- |
| **CAS — Hair & facial hair** | Scalp hair; facial hair; eyebrows; body hair; style/length/placement attributes where identifiable |
| **CAS — Clothing** | Tops; bottoms; full outfits; dresses; outerwear; swimwear; shoes; socks/hosiery; costumes; age/body-region and outfit-category facets |
| **CAS — Accessories** | Glasses; hats/headwear; earrings; necklaces; bracelets; rings; gloves; bags; nails; tattoos and other accessories (avoid forced subtype when uncertain) |
| **CAS — Makeup & facial detail** | Lipstick; eyeliner; eyeshadow; blush; face paint; eye masks/eyelids; beauty marks and other supported overlays |
| **CAS — Skin / Eyes / Body** | Skins and skin overlays; eyes/contacts; teeth; face and body presets or overlays; default replacements where applicable |
| **CAS — Sliders / Presets / Genetics** | Custom sliders, face/body presets and related genetic or CAS tuning resources; preserve actual technical content role |
| **CAS — Pets / Occults** | Species-specific or occult-specific CAS parts and accessories, with relevant compatibility metadata |
| **Buy Mode — Seating & Surfaces** | Chair/dining chair/office chair; armchair; sofa/loveseat; benches; tables/desks/counters/islands, etc. |
| **Buy Mode — Beds, Storage & Plumbing** | Single/double/bunk/crib; dresser/bookshelf/cabinet; sink/toilet/bath/shower and supported variants |
| **Buy Mode — Appliances & Electronics** | Refrigerator; stove; microwave; dishwasher; computers; TVs; stereos; phones; alarms, etc. |
| **Buy Mode — Lighting & Decorative** | Ceiling/wall/floor/table lights; wall art; sculptures; rugs; curtains; plants; clutter; decor objects |
| **Buy Mode — Skills, Entertainment & Special Objects** | Musical instruments; skill/hobby objects; gym; games; children's and pet objects; vehicles; utility, community-lot and special gameplay objects |
| **Build Mode — Openings & Structure** | Doors; arches; windows; fences/gates; stairs/railings; columns; foundations; platforms; roof objects |
| **Build Mode — Finishes & Landscaping** | Wall/floor coverings; terrain paints; ground cover; fireplaces/chimneys; pools/water; architectural details; trees, shrubs, flowers and gardening objects |
| **Patterns / Create-a-Style & Materials** | Patterns, textures, material swatches, recolours/overrides; distinguish content roles from placeable objects |
| **Gameplay / Scripts / Tuning** | Script mods, tuning changes, core/override resources, controllers and other non-catalogue content; not every item has a visual subtype |
| **Poses / Animations / Effects** | Pose sets, animations and effect resources where they can be identified safely; keep separate from CAS categories |
| **Lots / Sims / Worlds / Install Archives** | Lots, saved Sims, custom worlds and Sims3Pack/container contents, **where format-specific support is added**. Don't assume every archive is a single item |
| **Miscellaneous / Technical** | Defaults, fixes, overrides, resources without catalog entries, multi-purpose objects, unclassified or unknown items |

### Detection principles

- **Sims 3-specific parsers only.** Do not reuse Sims 2 OBJD/CTSS/XOBJ detection logic directly; Sims 3 uses different resources and catalogue systems.
- Prefer **verified resource metadata and relationships** (e.g. CASP for CAS; the relevant object and build resource families for objects) over filenames and folder names.
- Preserve **raw scanned evidence** (resource IDs, package/hash/path, flags, references) separately from the final human-readable classification.
- Respect **technical package roles**: new item vs recolour, override/default replacement, tuning, script or other content. Multiple resources do not automatically mean multiple distinct CC items.
- **One package can contain multiple distinct items**. Store individual item classifications linked to the parent package where feasible. A set distributed across several packages is not a merged collection.
- Allow **confirmed**, **inferred/suggested** and **unknown** values at the field level. No forced aesthetic subtypes: modern/traditional, materials, shape and other visual distinctions may need a thumbnail/preview or manual confirmation.
- **Unknown, invisible and intentionally technical content are not automatically broken.** Keep Health / Performance / Dependency / Classification separate.
- Preserve **manual changes and locks** across rescans. Existing classification data must remain valid until a versioned migration/backfill has been reviewed.

### Implementation approach — staged, non-disruptive

1. **Audit existing CCS3 classification.** Document which families, types and finer subtypes already work (especially CAS), which are missing, and which values are reliable. Avoid reimplementing completed features.
2. **Agree on a master taxonomy and field schema.** Shared cross-game concepts may be borrowed from CCS2's planning taxonomy, but keep game-specific semantics and Sims 3 detection rules.
3. **Map resource evidence → category/subtype.** Begin with high-confidence metadata-backed types, then add finer subtypes only after real-package validation. Record rules and confidence.
4. **Add optional library fields and filters.** Support Category / Subtype / source/confidence and relevant facets in Table, Gallery/detail preview, search and exports; respect column visibility and per-profile saved views.
5. **Add safe MDB migration and batch backfill.** Version the schema and detection rules, back up before altering historical records, allow re-analysis without rescanning unrelated data, and preserve manual overrides.
6. **Optional routing support later.** Allow users to organise by type/subtype only after preview/confirmation. Never move existing CC simply because a classification changes.
7. **Expand family-by-family.** CAS refinements first if useful, then supported Buy/Build object families, then non-object formats and complex cases as separate modules—never promise full automated visual recognition before testing.
8. **Regression checks and release planning.** Verify scanner speed, large libraries, duplicates, multi-item packages, archive handling, stable CAS results, filtering, storage persistence, profile sorting/routing and public-build compatibility.

### Examples

- A fireplace whose appearance cannot be determined remains **Fireplace**, not “Generic Fireplace”.
- A sofa whose seating subtype is verified can be **Seating → Sofa**.
- A CC face overlay can be **CAS → Makeup / Facial Detail → Eye Mask** where supported.
- A package with multiple resources for one CAS part is not a merged collection; several distinct independently identified CASPs/items require separate linked records.
- A tuning or script mod should be classified by its technical role even if it does not appear in CAS, Buy or Build catalogues.

### Deliverables / acceptance criteria

- A documented taxonomy covering existing CCS3 CC families plus explicit future families, with optional subtypes and no redundant Generic entries.
- A **capability matrix** for each family: supported now / parser needed / manual only / not tested.
- A tested dataset of representative real Sims 3 packages across all supported families, including ambiguous examples, multi-item packages, and default replacements.
- No regression in current CAS extraction, scan correctness, MDB read/write, backups, routing, public builds or profile-specific view state.
- No release announcement until tested; keep the future work separate from the current fix batch.

## Scheduling note

This is a **future CCS3 DEV enhancement**, not a demand to change the current stable public release, not a promise of a particular version, and not a request to modify CCS2. Resume this work after the active DEV scanner/routing reliability milestones have been completed.
