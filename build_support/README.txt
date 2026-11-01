CAS Curator v1.1.0 - Build Support Source
=========================================

These files document the build-only adapters used by the validated Windows release.
They are not imported by the normal source launcher.

The Windows release keeps S3PI/s3pe parsing behind a separate GPL helper process.
The parent scanner runtime does not bundle the S3PI DLLs. The Performance Score
Refresh fallback uses its own dedicated Python 3.8 worker runtime and the same
separate helper boundary.

The complete, tested one-click Windows build environment is distributed separately as:
CCS3_v1.1.0_WINDOWS_BUILD_KIT_FINAL.zip
