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

## Remaining gates before publication

1. Sign in to Netlify and record the currently PUBLISHED deployment ID/permalink and production branch. Match its source commit to the rollback reference; investigate any mismatch.
2. Confirm the previous deployment's rollback/publish action is available. Do not execute rollback during preparation.
3. Obtain a fresh read-only production backup immediately before publication. Existing full backup is Greenacre_Asset_Full_Backup_2026-10-22.zip, containing database JSON, manifest and all 30 attachments. Do not put client data, backups or credentials in Git.
4. Record current data version/counts and attachment IDs; earlier baseline was version 1424, 2 clients, 130 sites, 206 assets, 17 custom fields and 30 attachment files. Re-read, do not assume this baseline is still current.
5. Obtain Peter's approval to deploy the reviewed release. Preparing this branch/PR does not publish it.

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

No live rollback drill has been performed; that would change production. The source recovery point is created; exact Netlify deployment rollback verification awaits sign-in.
