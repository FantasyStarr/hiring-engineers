# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

This is `FantasyStarr/hiring-engineers`, a personal fork of DataDog's public [`hiring-engineers`](https://github.com/DataDog/hiring-engineers) take-home exercise. The upstream repo is a template: applicants fork it, complete the exercise on a role-specific branch (`support-engineer`, `solutions-engineer`, or `devops-evangelist` — see `README.md`), and submit a PR back upstream. This fork's `master` branch holds one candidate's completed write-up: tagging the Datadog Agent host, integrating a MySQL check, writing a custom Agent check, generating a Timeboard via the Datadog API, configuring a metric monitor with scheduled downtime, and collecting APM traces from a small Flask app. (The fork only has `master` — none of the upstream role-specific branches exist here; the work was committed directly to `master` instead of following the upstream fork-a-branch flow.)

**There is no application code, build system, package manifest, linter, test suite, or CI config in this repository.** Do not look for one, and do not invent `npm`/`make`/`pytest`-style commands — the entire deliverable is a Markdown write-up plus its supporting screenshots. Any code snippets you'll find (Python, YAML, HTML/Mustache) live *inside* `answers.md` as illustrative, already-submitted examples of what the candidate ran elsewhere — they aren't meant to be executed from this repo.

## Layout

- `README.md` — the unmodified upstream prompt/instructions from DataDog (what the exercise is, the three role branches, how to submit). Leave it as-is unless the user is deliberately updating the exercise prompt itself.
- `answers.md` — the actual submission and the file to edit for any content change. One long Markdown document walking through each exercise task in order, with narrative commentary, pasted config/code, and inline screenshots.
- `LICENSE` — Apache 2.0, carried over from the upstream template; it applies to the repo as forked, not meaningfully to the write-up content itself.
- Root-level `*.PNG` / `*.png` files — screenshots referenced from `answers.md`. There's no `images/`/`assets/` subfolder; everything is flat at repo root, and extensions are inconsistently `.PNG` vs `.png` (matters because GitHub raw URLs are case-sensitive).

## Conventions in `answers.md`

- Section headers use raw `<h2>`/`<h3>` HTML tags, not Markdown `##`/`###` — match this when adding a section.
- The file opens with a numbered Table of Contents that mirrors the `<h2>` sections below it; keep the two in sync when adding, renaming, or removing a section (they've already drifted slightly — e.g. the "Final Question" section isn't listed in the ToC).
- Each bonus question in the exercise gets its own `<h3>Bonus Question...</h3>` placed immediately after the main section it belongs to.
- Pasted scripts/config use fenced code blocks with a language hint (```python, ```yaml, ```html) matching the file being shown.
- Screenshots are embedded with `![screenshot](...)` right after the step they illustrate, using the **absolute** URL form `https://raw.githubusercontent.com/FantasyStarr/hiring-engineers/master/<filename>` — not a relative path. Two consequences:
  - A new screenshot must land as a flat file at repo root whose name matches (case included) the URL used in `answers.md`.
  - The link always resolves against `master`, regardless of what branch you're on — an image added only on a feature branch won't actually render in the write-up until it's merged to `master`.
- Prose is first-person and narrative ("I did X, then ran into Y", including a personal sign-off in the sample monitor email) — match that voice rather than switching to third-person documentation style.
- Several root-level images aren't referenced anywhere in `answers.md` (e.g. `apmworked.PNG`, `hostmaptag.PNG`, `metricsetup1.png`, `metricsetup2.png`, and five files literally named `Screenshot from 2017-12-05 ....png`). These look like extra/superseded captures kept from the original submission rather than dead weight — don't delete images as "unused cleanup" without checking with the user first, since nothing in the repo documents intent either way.

## Working in this repo

There's nothing to build, lint, or test. "Verifying" a change here means:
- The Markdown renders correctly (headers, fenced code blocks, image links).
- Any newly referenced screenshot filename actually exists at repo root, spelled and cased exactly as referenced.
- Commit messages follow the existing history's pattern of two kinds, usually kept separate: `Update answers.md` for text edits, and `Add files via upload` for screenshot additions.
