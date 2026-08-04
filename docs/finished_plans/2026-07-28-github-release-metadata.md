# GitHub Release and Repository Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish the sanitized current project state to `adadsws/r6s-stats-tiers` on `main`, then set the approved English description and GitHub topics.

**Architecture:** Keep the local source repository and its private history separate from the public history. Clone remote `main` into ignored `output/github-export/`, synchronize only explicitly publishable committed files, verify the exported tree and test suite, then perform a fast-forward-only push and update repository metadata through GitHub CLI.

**Tech Stack:** Git, GitHub CLI (`gh`), PowerShell, Python `unittest`

## Global Constraints

- Target repository is exactly `https://github.com/adadsws/r6s-stats-tiers`.
- Source branch, export branch, and remote default branch are all exactly `main`.
- Never force push and never create another remote branch or pull request.
- Publish ordinary source, tests, required documentation, sanitized sample input, and public reports.
- Exclude `output/`, `~temp/`, `~archived/`, `~ref/`, `docs/superpowers/`, non-example `secrets/`, `~$*.xlsx`, Git metadata, and agent runtime directories.
- Description is exactly `Generate Chinese Rainbow Six Siege operator tier lists and stats reports in XLSX and PDF.`
- Topics are exactly `rainbow-six-siege`, `r6s`, `operator-stats`, `tier-list`, `python`, `xlsx`, `pdf`, and `data-visualization`.
- Do not modify or commit unrelated local untracked content.

---

### Task 1: Establish the sanitized export repository

**Files:**
- Create locally and keep ignored: `output/github-export/`
- Read: `.gitignore`
- Read: `README.md`
- Read: `data/**`

**Interfaces:**
- Consumes: local committed `main` and remote `https://github.com/adadsws/r6s-stats-tiers`
- Produces: clean single-branch export checkout at `output/github-export/`

- [x] **Step 1: Verify source state and publication boundaries**

Run:

```powershell
git -c safe.directory=<PROJECT_ROOT> status --short
git -c safe.directory=<PROJECT_ROOT> branch --show-current
git -c safe.directory=<PROJECT_ROOT> branch --format='%(refname:short)'
git -c safe.directory=<PROJECT_ROOT> ls-files data secrets '~archived' '~ref' docs/superpowers
```

Expected: source branch is `main`; unrelated `~archived/output/` may remain untracked; the publication filter can account for every tracked private path.

- [x] **Step 2: Verify the export target is safe**

Resolve `output/github-export/` with PowerShell and confirm its absolute path starts with `<PROJECT_ROOT>\output\`. If an earlier export exists, preserve it by moving it to a timestamped sibling under `output/` before cloning.

- [x] **Step 3: Clone only remote main**

Run:

```powershell
git clone --branch main --single-branch https://github.com/adadsws/r6s-stats-tiers.git output/github-export
```

Expected: clone succeeds and `git -C output/github-export branch --format='%(refname:short)'` prints only `main`.

- [x] **Step 4: Synchronize the explicit publication allowlist**

Build the source list from committed files returned by `git ls-tree -r --name-only HEAD`. Exclude every path named in Global Constraints, allow only `secrets/**/*.example` under `secrets/`, and copy the remaining paths byte-for-byte into `output/github-export/`. Remove stale tracked export paths only when they are absent from this filtered committed source list.

- [x] **Step 5: Inspect the export diff**

Run:

```powershell
git -C output/github-export status --short
git -C output/github-export diff --check
git -C output/github-export diff --stat
```

Expected: only intended public project changes appear and no whitespace errors are reported.

### Task 2: Validate public contents and create the release commit

**Files:**
- Test: `output/github-export/tests/**`
- Inspect: every changed or newly added file in `output/github-export/`

**Interfaces:**
- Consumes: sanitized export worktree from Task 1
- Produces: one tested public commit based on remote `main`

- [x] **Step 1: Scan filenames and contents for private material**

List all exported files and changed paths. Search text files for credential patterns including `gho_`, `ghp_`, `github_pat_`, `AKIA`, `BEGIN PRIVATE KEY`, `api_key`, `access_token`, and absolute local user paths. Report only matching filenames, never secret values.

Expected: no credentials, private keys, private directories, or local absolute user paths are present.

- [x] **Step 2: Verify data is suitable for public release**

Inspect tracked `data/` filenames and schemas, confirm they contain Rainbow Six Siege public/game data and no personal identifiers, credentials, cookies, account tokens, or unrelated private input.

- [x] **Step 3: Run the full test suite in the export**

Run:

```powershell
$env:PYTHONDONTWRITEBYTECODE='1'
python -m unittest discover -s tests -v
```

Expected: all tests pass.

- [x] **Step 4: Run the pre-push hook when present**

If `.git/hooks/pre-push` exists in the export repository, invoke it with the same remote name and URL that the upcoming push will use. If no hook exists, record that explicitly.

- [x] **Step 5: Stage and review the exact public commit**

Run:

```powershell
git -C output/github-export add --all
git -C output/github-export diff --cached --check
git -C output/github-export diff --cached --stat
git -C output/github-export status --short
```

Expected: staged paths match the previously reviewed public diff.

- [x] **Step 6: Commit the public export**

Run:

```powershell
git -C output/github-export commit -m "feat: refresh operator report layouts"
```

Expected: one new public commit is created on export `main`.

### Task 3: Fast-forward push and repository metadata

**Files:**
- Modify remotely: GitHub repository `adadsws/r6s-stats-tiers`

**Interfaces:**
- Consumes: tested export commit from Task 2
- Produces: updated remote `main`, Description, and Topics

- [x] **Step 1: Re-fetch and enforce fast-forward publication**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export merge-base --is-ancestor origin/main main
```

Expected: exit code `0`; if it is not `0`, stop and synchronize remote changes before any push.

- [x] **Step 2: Push export main without force**

Run:

```powershell
git -C output/github-export push origin main:main
```

Expected: push succeeds as a fast-forward update.

- [x] **Step 3: Update Description and Topics**

Run:

```powershell
gh repo edit adadsws/r6s-stats-tiers `
  --description "Generate Chinese Rainbow Six Siege operator tier lists and stats reports in XLSX and PDF." `
  --add-topic rainbow-six-siege `
  --add-topic r6s `
  --add-topic operator-stats `
  --add-topic tier-list `
  --add-topic python `
  --add-topic xlsx `
  --add-topic pdf `
  --add-topic data-visualization
```

Remove any pre-existing topic not in the approved list so the resulting topic set is exact.

- [x] **Step 4: Verify remote content and metadata**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export rev-parse main
git -C output/github-export rev-parse origin/main
git -C output/github-export branch --format='%(refname:short)'
git -c safe.directory=<PROJECT_ROOT> branch --format='%(refname:short)'
gh repo view adadsws/r6s-stats-tiers --json description,repositoryTopics,defaultBranchRef,url
gh api repos/adadsws/r6s-stats-tiers/branches --paginate
```

Expected: export `main` equals `origin/main`; source and export each have only local `main`; the remote has only `main`; Description and Topics match Global Constraints.

- [x] **Step 5: Archive the completed plan locally**

Move this plan to `~archived/superpowers-plans/2026-07-28-github-release-metadata.md`, mark all steps complete, and commit the archive move only in the local source repository. Do not copy or push `docs/superpowers/` or `~archived/`.
