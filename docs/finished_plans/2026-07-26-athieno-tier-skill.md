# Athieno Tier Skill Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the project Skill to root `skills/` and restrict it to updating Athieno Tier data from the latest complete video.

**Architecture:** Keep the existing Skill invocation name while moving its self-contained folder. Enforce the path and single-responsibility boundary through the project layout test, then align README and UI metadata.

**Tech Stack:** Markdown, YAML, Python `unittest`, Skill Creator validation script, Git.

## Global Constraints

- Skill path is exactly `skills/build-r6-operator-report/`.
- Skill only writes `data/athieno/latest.json`.
- Skill does not collect Wiki data or generate XLSX, PDF, or previews.
- Do not use Microsoft Excel or LibreOffice.
- Do not push; create local commits only.

---

### Task 1: Lock the Skill contract with a failing test

**Files:**
- Modify: `tests/test_project_layout.py`
- Test: `tests/test_project_layout.py`

**Interfaces:**
- Consumes: repository root resolved from the test file.
- Produces: a regression contract for the root Skill path and single-purpose content.

- [x] **Step 1: Replace the old Skill layout test**

Assert that `skills/build-r6-operator-report/SKILL.md` exists, the old `.codex` path does not,
the new Skill contains `Athieno`, `YouTube`, `final_frame`, and
`data/athieno/latest.json`, and it does not contain `run_r6_report.bat`, `Wiki`, `.xlsx`,
`.pdf`, `Poppler`, `Microsoft Excel`, or `LibreOffice`.

- [x] **Step 2: Run the focused test and verify RED**

Run:

```powershell
$env:PYTHONPATH='src'
python -m unittest tests.test_project_layout.PackageLayoutTests.test_skill_only_updates_athieno_video_tier_data -v
```

Expected: fail because `skills/build-r6-operator-report/SKILL.md` does not exist.

### Task 2: Migrate and narrow the Skill

**Files:**
- Move: `.codex/skills/build-r6-operator-report/SKILL.md` to `skills/build-r6-operator-report/SKILL.md`
- Move: `.codex/skills/build-r6-operator-report/agents/openai.yaml` to `skills/build-r6-operator-report/agents/openai.yaml`

**Interfaces:**
- Consumes: the existing Athieno data contract.
- Produces: `$build-r6-operator-report`, whose only artifact is `data/athieno/latest.json`.

- [x] **Step 1: Rewrite `SKILL.md`**

Keep only official-video discovery, final-frame verification, complete operator membership,
fixed score mapping, JSON schema, safe overwrite rules, and JSON validation.

- [x] **Step 2: Rewrite `agents/openai.yaml`**

Use the display name `Athieno Tier 更新`, describe only video Tier extraction, and make the
default prompt explicitly invoke `$build-r6-operator-report`.

- [x] **Step 3: Validate GREEN**

Run the focused test again and run:

```powershell
python <USER_HOME>\.codex\skills\.system\skill-creator\scripts\quick_validate.py skills\build-r6-operator-report
```

Expected: both commands pass.

### Task 3: Align user documentation

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: the new root Skill path and single-purpose contract.
- Produces: local deployment steps that separate Tier acquisition from report generation.

- [x] **Step 1: Update README**

Change the Skill path to `skills/build-r6-operator-report/SKILL.md`. Keep Athieno acquisition
as deployment step 2, and remove report-generation instructions from the Skill prompt.

- [x] **Step 2: Update CHANGELOG**

Record the root-folder migration and single-responsibility scope.

### Task 4: Verify and archive

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-athieno-tier-skill.md` to `~archived/superpowers-plans/2026-07-26-athieno-tier-skill.md`

**Interfaces:**
- Consumes: all implementation changes.
- Produces: verified local commit with no active completed plan left under `docs/superpowers/plans/`.

- [x] **Step 1: Run full tests**

```powershell
$env:PYTHONPATH='src'
python -m unittest discover -s tests -v
```

Expected: all tests pass.

- [x] **Step 2: Check repository integrity**

Run `git diff --check`, confirm the root Skill is tracked and not ignored, and confirm the old
Skill path is absent.

- [x] **Step 3: Archive this plan and commit**

Move the completed plan into `~archived/superpowers-plans/`, stage all task files, and create
one local implementation commit without push.
