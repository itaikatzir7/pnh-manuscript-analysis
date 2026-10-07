# Private review to software DOI

Private review repository: [https://github.com/itaikatzir7/pnh-manuscript-analysis](https://github.com/itaikatzir7/pnh-manuscript-analysis). Do not treat this private review snapshot as v1.0.0.

## 1. Complete source and provenance

The three selected source files are included and hashed in source_inventory.json. Preserve source filenames and compute SHA-256 hashes with tools/check_repository.py. Compare newer candidates with the versions that produced the manuscript. The user-selected versions are authoritative for this package; no alternative programs are included.

## 2. Conservative cleanup

Capture original outputs before edits. Remove startup pip-install behavior from FRAP (completed in this package) in favor of explicit dependency instructions. Correct stale run commands and version labels only after source identity is established. Make output locations explicit and record all settings. Preserve equations, thresholds, bounds, normalization, statistics and exclusion rules.

For each edit, retain a short change log. Compare original and cleaned numeric results on the same input, settings and environment. Do not silently "fix" the FRAP half-time definition, trypsin sign, or plateau algorithm.

## 3. Validate

Use docs/VALIDATION.md in a new environment and a folder with spaces/non-ASCII characters. Record the full environment, actual settings and reference exports. Complete at least one end-to-end real example per included analysis. Record failures as failures.

## 4. Private repository and future updates

The repository was created privately on 7 October 2026. The steps below are retained for reference when preparing another repository; for this repository, clone it and commit subsequent changes.

1. Sign in to GitHub Desktop with the intended personal or lab account.
2. Choose File > New repository. Suggested name: pnh-manuscript-analysis.
3. Choose a local parent directory. Do not select an open-source license until the lab settles it.
4. Copy the completed repository CONTENTS into the newly created repository folder, without an extra nested project folder.
5. Review the Changes list. Exclude large unpublished raw data, temporary outputs and machine-specific paths.
6. Commit with a message such as "Prepare PNH manuscript analysis software for internal review".
7. Click Publish repository and leave **Keep this code private** checked. Confirm the intended account/owner.
8. Verify in the browser that the repository is marked Private.

The word "Publish" in GitHub Desktop does not require making the repository public. No invitations or messages have been sent by this preparation task.

## 5. Lab review

Have Ayala and relevant lab members review the numerical definitions, result reproducibility, contributors, software license and data permissions. Track corrections and commit them. Send invitations yourself when ready, or explicitly authorize them.

## 6. Public archive

After review and the chosen manuscript timing:
1. Confirm exact contributors, final software title, release version, license and repository URL in CITATION.cff.
2. Complete validation and compare manuscript numbers with frozen source/settings.
3. Change GitHub visibility to Public when appropriate.
4. Connect GitHub to Zenodo and enable THIS repository.
5. Only then create and publish GitHub release v1.0.0 from the reviewed commit. Enabling Zenodo before the release avoids relying on retrospective ingestion.
6. Verify that Zenodo archived the correct release and displays correct authors, files, version and license.
7. Cite the version-specific DOI in the manuscript. Record it in the repository documentation and commit the update without moving or rewriting the archived v1.0.0 tag.

Zenodo's GitHub integration requires a public repository. A private lab-review repository is not yet a DOI archive. Do not invent a DOI, ORCID, release date, contributor or journal citation.

## 7. Availability statement template

Use only after the repository and record exist:

"Custom analysis software used in this study is available at [REPOSITORY URL]. The version used for the reported analyses is archived on Zenodo as version 1.0.0 ([VERSION-SPECIFIC DOI]). Analysis settings, representative input data, and reproduction instructions are provided with the software. [Describe the actual location and access conditions of the complete datasets.]"

Avoid claiming that this repository reproduces every figure unless all manual/commercial-software steps and data are covered.

## Official instructions checked on 6 October 2026

- [GitHub Desktop: first repository](https://docs.github.com/en/desktop/overview/creating-your-first-repository-using-github-desktop)
- [GitHub: add an existing project](https://docs.github.com/en/desktop/adding-and-cloning-repositories/adding-an-existing-project-to-github-using-github-desktop)
- [GitHub: citing and archiving](https://docs.github.com/en/repositories/archiving-a-github-repository/referencing-and-citing-content)
- [Zenodo: enable a repository](https://help.zenodo.org/docs/github/enable-repository/)
- [Zenodo: release archival](https://help.zenodo.org/docs/github/archive-software/github-upload/)
- [Zenodo: software metadata](https://help.zenodo.org/docs/github/describe-software/)

Use CITATION.cff as the single metadata source initially; a .zenodo.json, if added, takes precedence in Zenodo and must be maintained consistently.
