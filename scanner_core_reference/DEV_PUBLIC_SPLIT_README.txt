Sims 3 CC Health Scanner — DEV / PUBLIC shared-core split
=========================================================

This build introduces two launchers without duplicating scanner logic:

  Sims3_CC_Health_Scanner_DEV.py
  Sims3_CC_Health_Scanner_PUBLIC.py

Both launch:

  Sims3_CC_Health_Scanner_CORE.py

Important guarantees
--------------------
- DEV and PUBLIC use the same health checks, classifiers, routing logic and MDB writer.
- The permanent Master Database keeps the FULL field set in both editions.
- PUBLIC does not create a reduced MDB schema. A future PUBLIC GUI can simply hide
  most columns by default while still exposing 100% of fields through column selection
  and the package detail panel.
- PUBLIC forcibly disables developer debug switches.
- DEV keeps debug switches available but OFF by default, preserving the current quiet
  validation behaviour.
- This split does not add CC Magic-style set/merge/profile management.

DEV-only environment switches
-----------------------------
SIMS3_CC_DEBUG=1
SIMS3_CC_DEBUG_RESOURCES=1
SIMS3_CC_DEBUG_BUILD=1
SIMS3_CC_EXPERIMENTAL=1
SIMS3_CC_QUIET=0

PUBLIC ignores the debug/experimental switches and remains in quiet mode.

Current scope
-------------
This is architecture preparation for the eventual EXE. It does not build the GUI/EXE yet.
The existing Scan_Summary.xlsx design and permanent MDB behaviour are preserved.


2026-09-07 EDITION LOG FIX TEST
-------------------------------
Quiet logs now retain the current shared-core BUILD marker and the
SCANNER EDITION line. This makes DEV vs PUBLIC identifiable from an
uploaded Health_Scan_Log.txt. Scanner logic, MDB schema, routing, and
health checks are unchanged.

2026-09-07 DEFAULT REPLACEMENT AUXILIARY PREVIEW FIX TEST
- Dedicated preview/icon/thumbnail resource types no longer prevent an otherwise proven modified EA-resource package from becoming EA Override.
- This specifically fixes true default replacements that contain modified EA TGIs plus creator-added custom preview images.
- Packages with zero EA TGI overlap remain Custom Content; filename wording alone never makes a package a default replacement.
- No health-check functions were removed or changed by this patch.


2026-09-07 TEST - Script Mod with embedded CAS primary classification
- A valid S3SA + tuning/ITUN package that bundles CASP outfit resources can now
  classify primarily as Script Mod when it contains no Build/Buy object resources.
- Embedded CASP records remain in CASP Master with their individual CAS types.
- This prevents service/script mods such as douglasveiga_GardenerService2.package
  from being routed primarily as Mixed CAS.
- Ordinary CAS/Mixed CAS, Custom Sims and scripted Build/Buy objects are unchanged.


2026-09-08 TEST - Scripted Object taxonomy + subtype
---------------------------------------------------
- Primary type is now `Scripted Object` rather than `Scripted Object Mod`.
- `Embedded Script` subtype: object package contains its own valid S3SA.
- `External Script` subtype: no S3SA in the object package, but a healthy OBJK
  points to conservatively detected custom external script-class evidence.
- Embedded-script object detection no longer requires tuning/ITUN in addition to S3SA.
- Ordinary Build/Buy objects that only reference normal EA classes remain Build/Buy.
- Pattern remains protected from generic Build/Buy override.
- Structural subtype is written to the existing MDB Subtype field and displayed as
  `Scripted Object — Embedded Script` or `Scripted Object — External Script`.
- All 66 existing check_* health checks are preserved.


2026-09-12 SCRIPTED OBJECT EVIDENCE REFINEMENT TEST
- Embedded Scripted Object now requires valid S3SA + OBJD + an OBJK scriptClass rooted in Sims3.Gameplay.Objects.*.
- This retains PascalCase creator namespaces such as Dexter / TS3Apartments.
- External Scripted Object keeps the stricter custom-class heuristic and now also requires OBJD.
- S3SA packages with object resources that do not meet embedded-object evidence fall back to Script Mod.
- Quiet logs now reliably retain the current PRIMARY_BUILD_MARKER.
- No health check logic changed.


2026-09-12 SCRIPTCLASS TYPEDEF MATCH TEST
- Embedded Scripted Object now requires OBJK scriptClass == a .NET TypeDef actually declared in local valid S3SA.
- External Script ignores lowercase leaf class names and requires a creator-like lowercase namespace segment before the leaf.
- Purpose: keep Ani/true scripted objects, stop Icolopolis ordinary-EA-class false positives, and recover locally-defined scripted objects without filename rules.
- TEST only until Windows/S3PI regression validation.


2026-09-12 SCRIPTCLASS TYPEDEF MATCH PERMANENT
------------------------------------------------
Promoted unchanged from the validated SCRIPTCLASS TYPEDEF MATCH TEST after the
345-package Windows/S3PI Big Review regression scan completed successfully.

Validated behaviour includes:
- Embedded Scripted Object requires an OBJK scriptClass matched to a .NET TypeDef
  actually declared in the package's valid S3SA, with the package-level guards
  used by the validated TEST build.
- External Scripted Object requires conservative creator-namespace evidence and
  does not treat a lowercase leaf EA class name as creator-script evidence.
- ani_BistroMenu / ani_Chicken remain External Script.
- Ani embedded-script objects, Granthes OMSP, Butterbot Canning Station,
  Mingos Birds & Bees, cmomoney Cigarette Mod and other validated examples remain
  Embedded Scripted Object.
- NRaas_Career and NRaas_PocketProtectorTablet remain Script Mod rather than being
  promoted merely because they contain script/object resources.
- ICOLOPOLIS_GrandeLama_newCC no longer false-positives as Scripted Object.
- Merged packages that genuinely contain a locally-defined scripted object may
  still classify as Scripted Object (for example ICOLOPOLIS_variousCC containing
  the Granthes OMSP TypeDef).
- No health-check logic was changed for this promotion; all 66 check_* functions
  are preserved.
