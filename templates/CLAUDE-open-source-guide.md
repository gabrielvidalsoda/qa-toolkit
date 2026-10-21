# Open Source Project Guide <!-- template v1.0 -->

> **How to use this file:** copy it into a new repo as `.claude/CLAUDE.md` (or
> paste its content into the project's existing `CLAUDE.md`). Fill in the
> `{{PLACEHOLDERS}}` at the top, delete any section that doesn't apply, then
> let Claude Code follow it automatically for that project.
> Based on [opensource.guide/best-practices](https://opensource.guide/best-practices/)
> and [10up's Open Source Best Practices](https://10up.github.io/Open-Source-Best-Practices/).

## Project info (fill in per project)

- **Maintainer:** {{YOUR_NAME}}
- **Contact email:** {{YOUR_EMAIL}}
- **Other contact (optional):** {{PHONE / WHATSAPP LINK / LINKEDIN / DISCORD}}
- **License:** {{MIT / Apache-2.0 / GPL-3.0 / ...}}
- **Primary doc language:** {{English / other}} — **{{secondary language, or "none"}}** as a secondary translated doc, if any
- **Code/CLI language:** always English, regardless of doc language above

## Baseline: files every repo should have

When working in this project, keep these present and up to date. If one is
missing and the task touches related code, create it rather than skipping it:

- `README.md` — what the project does, why it exists, install/usage, and a
  **Contact** section with the info above. Add badges only if they're true
  (license, CI status, Python/Node version) — no aspirational badges.
- `LICENSE` — matching the license chosen above. Don't leave a public repo
  unlicensed; ask the maintainer once if it's genuinely undecided, then don't
  ask again for the rest of the session.
- `CONTRIBUTING.md` — how to propose a change, dev setup, code style, how
  changes are tested, expected response time, link to Code of Conduct.
- `CODE_OF_CONDUCT.md` — Contributor Covenant is a safe default; put the
  maintainer's contact email in the enforcement section.
- `.github/SECURITY.md` (or root `SECURITY.md`) — how to privately report a
  vulnerability (email, not a public issue), and what's in/out of scope.
- `CHANGELOG.md` — Keep a Changelog format, `Unreleased` section at the top;
  add an entry for every user-facing change, not just releases.
- `.github/ISSUE_TEMPLATE/` (bug report + feature request) and
  `.github/PULL_REQUEST_TEMPLATE.md` — lowers the barrier for outside
  contributors and keeps low-quality/low-context reports rarer.
- A minimal CI workflow (`.github/workflows/`) — at least syntax check +
  lint. Don't add a full test matrix speculatively if there are no tests yet;
  add real tests only when the task calls for them.

## Language strategy for bilingual maintainers

If the maintainer's first language isn't English but they want the project
usable by an international audience:

- **Code, comments, docstrings, CLI/log output, commit messages:** English,
  always. This is non-negotiable — it's what makes the project reviewable and
  usable by contributors who don't share the maintainer's first language.
- **Docs (README, CONTRIBUTING, etc.):** English is the primary/canonical
  version. A translated version is a nice-to-have, named `README.pt-BR.md`
  (or the appropriate ISO code), cross-linked from both files with a language
  switcher at the top (`🇬🇧 English | 🇧🇷 [Português](README.pt-BR.md)`).
- If a translated doc gets out of sync with English during a change, it's
  fine to update English first and flag the translation as pending — don't
  block a merge on the translation unless the maintainer says otherwise.

## Contribution workflow expectations

- Prefer small, focused PRs over large ones.
- Don't accept/design for a feature that expands scope beyond what the
  README already promises — flag scope creep instead of silently building it.
- Every behavior/flag/config-format change updates the README (and its
  translation, if any) and gets a `CHANGELOG.md` entry under `Unreleased`.
- Never commit secrets, tokens, `.env` files, or real customer/client data —
  check `.gitignore` covers them before the first commit that could include
  them, not after.
- Example/template config files (`*.example.*`) must contain only generic,
  clearly-fake placeholder values — never a real client name, slug, URL, or
  credential copied from a working local config.

## Commit and PR hygiene

- Conventional, descriptive commit messages (what changed and why, not just
  "fix" or "update").
- PRs describe: what changed, why, how it was tested, and link the issue if
  one exists.
- Don't force-push shared branches or rewrite published history without the
  maintainer's explicit go-ahead.

## What NOT to do

- Don't invent a Code of Conduct enforcement contact other than the
  maintainer's own listed contact.
- Don't add a license the maintainer hasn't confirmed for that repo — ask
  once via a direct question if it's genuinely unset.
- Don't bilingual-ize the codebase itself (no parallel `_pt.py` files) —
  only documentation gets a translated counterpart.
- Don't scaffold a large test suite, CI matrix, or release-automation
  pipeline speculatively for a small single-maintainer utility — match the
  tooling to the project's actual size and maturity.
