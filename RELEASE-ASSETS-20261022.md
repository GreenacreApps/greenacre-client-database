# Asset enhancement release — 22 October 2026

Status: prepared, not merged or deployed.

## Release references

- Repository: GreenacreApps/greenacre-client-database
- Release branch: release/assets-enhancement-20261022
- Rollback branch: rollback/assets-before-enhancement-20261022
- Pinned rollback commit: 76aff3098540c08745ea5b78e59868ef63d727e8
- Production Netlify project: greenacreclientassetdatabase
- Production URL: https://greenacreclientassetdatabase.netlify.app/
- Isolated test project: greenacre-assets-test-20261022

## What ships

Responsive registers, unified search and breadcrumbs; grouped asset details; site/asset thumbnails and asset galleries; saved cover selection when multiple photos exist; strict PDF exclusion; contact fields excluded from visible screens and Excel import/export; one visible Project Numbers field preserving existing stored values.

The release starts from production main. app.js and styles.css match the tested versions. index.html has the test banner removed. The production API differs only in the two contact-import exclusion checks. Production authentication, storage selection, reminder routes, weekly schedule, dependencies and Netlify configuration are retained. No test restore endpoint, transfer page or responsive harness is included.

No database migration, initial seed, overwrite, attachment move or contact-history deletion is performed. coverPhotoId and projectNumbersUnified are optional asset properties saved only through user edits. Rollback code retains unknown asset properties via existing serialization.

## Validation completed

- Browser layouts: 390, 768, 1024 and 1366 CSS pixels; no page overflow/clipped panels.
- Navigation, search/empty states, attachment photos, cover selection/reload, and mobile dialogs.
- Reversible edit persisted after reload and original value restored.
- PDF-only assets show no image or cover selector; PDF links retained.
- Excel export: 2 clients, 130 sites, 206 assets, one Project Numbers column, no contact fields.
- Test record/custom field comparison against backup passed after restoring temporary test edits.
- Peter confirmed actual phone/tablet touch, pinch zoom and camera upload work on 22 October 2026.
- Release-specific checks: node scripts/check-production-release.mjs (Node 24 used).
- Syntax: node --check app.js; node --experimental-strip-types --check netlify/functions/api.mts.

The automated release test uses synthetic in-memory stores only. It checks production/preview storage selection, authentication, version conflicts, lossless data roundtrip, contact-import exclusion, project values, selected-cover fallback and PDF exclusion.

## Production preparation verified — 23 October 2026

- Netlify sign-in completed. Published production is main@76aff30, matching the pinned rollback source.
- Rollback deployment: 6a994e4de8ac3d0008ad266e.
- Deployment details: https://app.netlify.com/projects/greenacreclientassetdatabase/deploys/6a994e4de8ac3d0008ad266e
- Immutable permalink: https://6a994e4de8ac3d0008ad266e--greenacreclientassetdatabase.netlify.app
- The same project's older completed production deploy exposes Publish deploy. The recovery control was inspected without clicking it. No live rollback drill was performed.
- Release code at 6c39f51a52607c6b4e3976e3fad9e2ae7c002708 built successfully as deploy preview 6ada1f464fcbeb19c8957ca6 (12 seconds).
- Fresh read-only database backup: Greenacre_Asset_PreRelease_2026-10-23_R1425.json, 224752 bytes; version 1425; 2 clients, 130 sites, 206 assets, 17 custom fields. Saved separately from Git.
- All 30 current live attachment IDs across 16 asset folders exactly match Greenacre_Asset_Full_Backup_2026-10-22.zip. All 30 archived file SHA-256 checks passed. Retain that full backup together with the fresh database JSON.
- No production code or data was changed during preparation.

## Remaining gate before publication

Obtain Peter's final approval to deploy this reviewed release. Immediately before publishing, recheck the published baseline and data version; refresh the backup if legitimate user edits occurred. Preparation does not publish the release.

## Deployment after approval

- Confirm main and the published deployment are still the recorded baseline.
- Merge the reviewed release into main; allow the existing production site's main deployment to build.
- Use the existing production project, storage and environment variables. Never promote the isolated test site's deploy or copy its store into production.
- Verify successful build/published commit, /api/health, sign-in, actual counts/IDs, original attachments, desktop/mobile layouts and cover selection.
- Do not send test emails. Existing production reminders retain their prior configuration.
- Compare production records with the fresh backup, allowing only explicitly authorized user changes/version metadata.
- If any data, storage, sign-in, attachment or core navigation check fails, stop edits and perform code rollback.

## Code rollback

Preferred: in the SAME production Netlify project, select the previously recorded published deployment and use its available rollback/publish action. Verify the old deployment becomes published, then verify sign-in, original counts/IDs and attachments.

Also reconcile main with the rollback source through a reviewed revert PR so the next automatic build does not reintroduce the failed release. Do not force-push main.

Fallback if the old deploy is unavailable: create a revert release restoring the affected production files from pinned commit 76aff3098540c08745ea5b78e59868ef63d727e8, then build/publish through the normal production project.

Code rollback retains the current data store and all edits/uploads made since deployment. It is NOT a backup restore. Do not restore the pre-release database by default, because that could erase legitimate subsequent edits. A data restore needs a separate reconciliation and explicit approval.

No live rollback drill has been performed; that would change production. The source recovery point and exact published Netlify deployment are recorded above; the Publish deploy recovery control was verified on an older completed deploy.
